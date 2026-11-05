"""Startup-time tool filtering via the POWERPOINT_MCP_TOOLSETS environment variable.

Mirrors word-mcp-live's WORD_MCP_TOOLSETS and excel-mcp-live's
EXCEL_MCP_TOOLSETS: a comma-separated union of groups, where unset, empty or
``all`` registers every tool. ``list_open_presentations`` is always
registered, since every other tool needs a presentation name to target.
"""

from fastmcp.tools import Tool

from .tool_metadata import drop_null_defaults, prepare_tool


ENV_VAR = "POWERPOINT_MCP_TOOLSETS"

# Always registered, whatever the selection.
CORE_TOOLS = frozenset({"list_open_presentations"})

TOOL_GROUPS = {
    "presentations": (
        "list_open_presentations", "get_presentation_info", "save_presentation",
    ),
    "slides": (
        "list_slides", "add_slide", "delete_slide", "duplicate_slide", "move_slide",
        "activate_slide",
    ),
    "shapes": (
        "list_shapes", "get_slide_text", "add_textbox", "set_shape_text", "delete_shape",
        "set_shape_position",
    ),
    "formatting": (
        "format_shape_text", "format_shape",
    ),
    "notes": (
        "get_speaker_notes", "set_speaker_notes",
    ),
    "export": (
        "export_slide_to_image",
    ),
    "comments": (
        "add_comment", "get_comments", "delete_comment",
    ),
}

TOOLSETS = frozenset(TOOL_GROUPS)

_GROUP_OF = {name: group for group, names in TOOL_GROUPS.items() for name in names}


def tool_group(name):
    """The toolset a tool belongs to. Every registered tool must have one."""
    try:
        return _GROUP_OF[name]
    except KeyError:
        raise KeyError(f"Tool {name!r} is not assigned to a toolset in toolsets.TOOL_GROUPS") from None


def parse_toolsets(value):
    """Parse a comma-separated union of toolsets; None means all tools."""
    if value is None or not value.strip():
        return None
    parts = [part.strip().lower() for part in value.split(",")]
    if any(not part for part in parts):
        raise ValueError(f"{ENV_VAR} contains an empty toolset name")
    if parts == ["all"]:
        return None
    unknown = set(parts) - TOOLSETS
    if unknown:
        raise ValueError(
            f"Unknown {ENV_VAR} value(s): " + ", ".join(sorted(unknown))
            + ". Choose from all, " + ", ".join(sorted(TOOLSETS))
        )
    return frozenset(parts)


class ToolRegistrar:
    """Stand-in for ``FastMCP`` in register_tools(): filters by toolset and compacts metadata."""

    def __init__(self, server, selected):
        self.server = server
        self.selected = selected

    def tool(self, **kwargs):
        def register(function):
            name = function.__name__
            group = tool_group(name)
            if self.selected is None or name in CORE_TOOLS or group in self.selected:
                prepared, options = prepare_tool(function, **kwargs)
                tool = Tool.from_function(prepared, **options)
                drop_null_defaults(tool.parameters)
                self.server.add_tool(tool)
            return function
        return register
