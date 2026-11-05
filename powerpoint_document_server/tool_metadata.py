"""Compact MCP tool metadata built from the tool modules' docstrings.

The docstrings in ``tools/*.py`` stay the single source of truth. At
registration time, ``prepare_tool()`` splits a Google-style docstring into:

* a short tool description: the prose before ``Args:``, plus anything in
  ``Returns:`` that says more than "JSON confirmation";
* per-parameter descriptions, attached to the input schema with
  ``Annotated[..., Field(description=...)]`` so clients show them next to
  each parameter instead of as prose in the description.

Descriptions of the shared ``presentation`` / ``slide_index`` / ``shape``
parameters that only restate the server-wide convention (see
``SERVER_INSTRUCTIONS``) are dropped rather than repeated on every tool.
Ported from excel-mcp-live's ``tool_metadata.py``; keep the two in step.
"""

from __future__ import annotations

import functools
import inspect
import re
from typing import Annotated, Any, Callable, get_args, get_origin, get_type_hints

from pydantic import Field
from pydantic.fields import FieldInfo


SERVER_INSTRUCTIONS = (
    "Edits presentations already open in Microsoft PowerPoint (Windows). "
    "Parameters shared by most tools: `presentation` is an open presentation's name or full "
    "path, and omitting it means the active presentation; `slide_index` is a 1-based slide "
    "index (required); `shape` is a shape's name or its 1-based index on the slide "
    "(required), as listed by list_shapes. Positions and sizes are in points; colors are hex "
    "RGB such as \"FF0000\" (a leading '#' is allowed). Parameters marked (required) have no "
    "usable default. Every tool returns a JSON string; failures return {\"error\": ...}."
)

# Parameter descriptions that only restate SERVER_INSTRUCTIONS. Anything more
# specific (e.g. the slide tools' `index`, or delete_comment's comment index)
# is kept.
_BOILERPLATE = {
    "presentation": re.compile(r"Presentation name or path \(None = active presentation\)\."),
    "slide_index": re.compile(r"1-based slide index \(required\)\."),
    "shape": re.compile(r"Shape name, or its 1-based index on the slide \(required\)\."),
}

_SECTION = re.compile(r"^(Args|Arguments|Returns):\s*$")
_ARG = re.compile(r"^(\s+)(\w+(?:\s*,\s*\w+)*):\s*(.*)$")
_BARE_CONFIRMATION = re.compile(r"^JSON confirmation(?: of [^.]*)?\.?\s*", re.IGNORECASE)


def _join(lines: list[str]) -> str:
    return " ".join(line.strip() for line in lines if line.strip())


def parse_docstring(doc: str) -> tuple[str, dict[str, str], str]:
    """Split a Google-style docstring into (prose, {param: description}, returns)."""
    prose: list[str] = []
    returns: list[str] = []
    args: dict[str, list[str]] = {}
    section = "prose"
    current: list[str] = []
    arg_indent: int | None = None

    for line in inspect.cleandoc(doc or "").splitlines():
        heading = _SECTION.match(line)
        if heading:
            section = "returns" if heading.group(1) == "Returns" else "args"
            arg_indent = None
            continue
        if section == "prose":
            prose.append(line)
        elif section == "returns":
            returns.append(line)
        else:
            match = _ARG.match(line)
            if match and (arg_indent is None or len(match.group(1)) == arg_indent):
                arg_indent = len(match.group(1))
                # "bold, italic, underline: ..." documents several parameters at once.
                current = [match.group(3)]
                for name in re.split(r"\s*,\s*", match.group(2)):
                    args[name] = current
            elif line.strip():
                current.append(line)

    return _join(prose), {name: _join(text) for name, text in args.items()}, _join(returns)


def compact_description(prose: str, returns: str) -> str:
    """Tool summary plus whatever the Returns section adds beyond a bare confirmation."""
    extra = _BARE_CONFIRMATION.sub("", returns, count=1)
    if extra.startswith("JSON with "):
        extra = "Returns " + extra[len("JSON with "):]
    elif extra.startswith("with "):
        extra = "Returns " + extra[len("with "):]
    return " ".join(part for part in (prose, extra) if part)


def _is_boilerplate(name: str, description: str) -> bool:
    pattern = _BOILERPLATE.get(name)
    return bool(pattern and pattern.fullmatch(" ".join(description.split())))


def drop_null_defaults(schema: dict[str, Any]) -> dict[str, Any]:
    """Remove ``"default": null`` from top-level properties, in place.

    An omitted optional parameter already means None, so the explicit null
    default on every optional parameter is pure bulk in tools/list.
    """
    for prop in schema.get("properties", {}).values():
        if "default" in prop and prop["default"] is None:
            del prop["default"]
    return schema


def _has_field(annotation: Any, default: Any) -> bool:
    if isinstance(default, FieldInfo):
        return True
    return get_origin(annotation) is Annotated and any(
        isinstance(item, FieldInfo) for item in get_args(annotation)[1:]
    )


def prepare_tool(function: Callable[..., Any], **registration_kwargs: Any) -> tuple[Callable[..., Any], dict[str, Any]]:
    """Return a wrapper for ``function`` and FastMCP registration kwargs with compact metadata.

    The docstring is taken from ``description=`` if given (main.py passes the
    tool module's docstring), else from ``function.__doc__``. Parameter names,
    defaults and types are unchanged; only schema descriptions are added.
    Existing ``Field`` metadata is left alone.
    """
    doc = registration_kwargs.get("description")
    if doc is None:
        doc = function.__doc__ or ""
    prose, arg_docs, returns = parse_docstring(doc)
    description = compact_description(prose, returns)

    unknown = set(arg_docs) - set(inspect.signature(function).parameters)
    if unknown:
        raise ValueError(
            f"{function.__name__}: docstring documents unknown parameter(s) {sorted(unknown)}"
        )

    signature = inspect.signature(function)
    hints = get_type_hints(function, include_extras=True)
    parameters = []
    for parameter in signature.parameters.values():
        hint = hints.get(parameter.name, parameter.annotation)
        arg_doc = arg_docs.get(parameter.name, "")
        if (arg_doc and hint is not inspect.Parameter.empty
                and not _is_boilerplate(parameter.name, arg_doc)
                and not _has_field(hint, parameter.default)):
            hint = Annotated[hint, Field(description=arg_doc)]
            hints[parameter.name] = hint
            parameter = parameter.replace(annotation=hint)
        parameters.append(parameter)

    @functools.wraps(function)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        return function(*args, **kwargs)

    wrapper.__doc__ = description
    wrapper.__annotations__ = hints
    wrapper.__signature__ = signature.replace(parameters=parameters)  # type: ignore[attr-defined]
    kwargs = dict(registration_kwargs)
    kwargs["description"] = description
    return wrapper, kwargs
