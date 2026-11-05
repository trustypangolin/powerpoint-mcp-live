# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Goal

`powerpoint-mcp-live` gives an AI assistant control of a PowerPoint presentation the **user**
already has open — no separate open/close lifecycle, no server-managed PowerPoint process. It's the
PowerPoint counterpart to [`../word-mcp-live`](../word-mcp-live)'s live COM tools and
[`../excel-mcp-live`](../excel-mcp-live), built the same way: by porting `word-mcp-live`'s attach
pattern (`word_document_server/core/word_com.py`) to a new COM target.

This exists because the workspace's other PowerPoint automation project,
[`../mcp-server-powerpoint`](../mcp-server-powerpoint), works the opposite way by design: it
launches its own `PowerPoint.Application` process, owns it for the life of a session, and closes it
on shutdown — the right model for scripted, from-scratch deck generation, but it can't attach to a
deck the user is looking at. `powerpoint-mcp-live` fills that specific gap: attach to what's already
open, edit it live, never touch its process lifetime. Use `mcp-server-powerpoint` for tables,
charts, SmartArt, animations, templates/themes, and anything that should run start-to-finish
without a human's PowerPoint window in the loop; use `powerpoint-mcp-live` when the user has a
presentation open and wants the assistant to edit *that* window.

## Non-negotiable design constraint

**This server must never own a PowerPoint process.** Every tool attaches via
`powerpoint_document_server/core/powerpoint_com.py`'s `get_powerpoint_app()` /
`find_presentation()`, which use `win32com.client.GetActiveObject("PowerPoint.Application")` plus a
Running Object Table (ROT) scan — the same pattern as `word_com.py`'s `get_word_app()` /
`_find_word_with_docs()`. No tool may call `Application.Quit()`, and no tool may close a
presentation the caller didn't explicitly ask to close (there isn't a "close" tool yet — see
Roadmap). If you ever add a "start PowerPoint" or "create presentation" tool here, it does not
belong in this project — that's `mcp-server-powerpoint`'s job.

**Design note — why this is closer to word_com.py than excel_com.py:** PowerPoint, like Word, is
effectively single-instance: opening multiple presentations opens multiple windows within one
`PowerPoint.Application` process, not separate processes (unlike Excel, which commonly runs several
simultaneous `Excel.Application` processes). So `powerpoint_com.py` mirrors `word_com.py`'s
single-app `get_powerpoint_app()` pattern rather than `excel_com.py`'s multi-instance
`list_all_open_workbooks()` ROT walk. If real-world testing shows PowerPoint does end up with
multiple processes in some configuration, port the multi-instance pattern over from `excel_com.py`.

## Known limitation: no undo grouping

Word's COM API exposes `Application.UndoRecord.StartCustomRecord`/`EndCustomRecord`, letting
`word_com.py`'s `undo_record()` context manager collapse a multi-step edit into a single Ctrl+Z.
**PowerPoint has no equivalent**, same as Excel (see `../excel-mcp-live/CLAUDE.md`). A tool that
performs several COM writes to accomplish one logical edit (e.g. adding a text box and then styling
it) will require multiple Ctrl+Z's to fully undo, not one. Don't claim single-undo behavior in tool
descriptions or the README.

## Commands

```powershell
# Install (editable, for local development)
pip install -e ".[dev]"

# Run the server directly against source (no install needed)
uv run powerpoint_mcp_server.py

# Syntax-check everything
python -m py_compile powerpoint_document_server/**/*.py

# tools/list shape, payload budget and toolset tests (no PowerPoint needed)
uv run --extra dev pytest
```

`tests/test_tools_list.py` covers registration only: descriptions, schemas, hints, payload size
and `POWERPOINT_MCP_TOOLSETS` filtering. It never touches COM. The tools' behavior has no automated
tests. Like `mcp-server-powerpoint` and `mcp-server-excel`, COM automation isn't meaningfully
mockable — real verification means opening PowerPoint with a presentation and exercising tools
against it by hand.

## Architecture

```
powerpoint_document_server/
  main.py               FastMCP server — every tool registered here, description=impl.__doc__
                         pulls each tool's docstring in (single source of truth — write real
                         docstrings, don't duplicate them in main.py)
  tool_metadata.py       Turns those Google-style docstrings into compact metadata: the prose
                         (plus any Returns: caveat) becomes the description, and each Args: entry
                         becomes that parameter's schema description. Boilerplate presentation/
                         slide_index/shape docs are dropped; SERVER_INSTRUCTIONS states them once.
                         A docstring Args: entry that names no real parameter fails at startup.
                         Ported from excel-mcp-live — keep the two in step.
  toolsets.py            POWERPOINT_MCP_TOOLSETS filtering. Every new tool must be added to exactly
                         one group in TOOL_GROUPS, or registration raises KeyError.
  defaults.py            MCP_AUTHOR / MCP_AUTHOR_INITIALS env vars → comment author defaults
  core/
    powerpoint_com.py     COM attach layer: get_powerpoint_app(), find_presentation(),
                           find_slide(), find_shape() — the only place GetActiveObject/ROT-scanning
                           happens
  utils/
    conversions.py        hex_to_bgr() (PowerPoint's ColorFormat.RGB uses BGR int, not RGB hex —
                           same convention as Excel/Word)
  tools/
    presentation_tools.py  list/inspect/save open presentations
    slide_tools.py          add/delete/duplicate/move/activate slides
    shape_tools.py          list shapes, read/write text, add text boxes, position/delete shapes
    format_tools.py         font style/color, shape fill/outline
    notes_tools.py          speaker notes (NotesPage.Shapes.Placeholders(2))
    export_tools.py         export a slide to an image (export-to-verify, matching
                             mcp-server-powerpoint's differentiator — see its README)
    comment_tools.py        slide comments (PowerPoint 2016+ modern comments)
```

Every tool function follows the same shape: guard on `sys.platform == "win32"`, lazily import from
`core.powerpoint_com` inside the function body (keeps the module importable on non-Windows), call
`find_presentation()`/`find_slide()`/`find_shape()` to attach, do the COM operation, return a JSON
string (`json.dumps({...})` on success, `json.dumps({"error": ...})` on failure — tools never
raise).

## Roadmap / not yet implemented

- Tables, charts, SmartArt, animations, transitions, templates/themes, slide masters — intentionally
  out of scope; that's `mcp-server-powerpoint`.
- No explicit "close presentation" or "detach" tool.
- No macOS support.
- No automated tests of tool behavior (only of registration, see Commands).
- `export_slide_to_image` and the comment tools are unverified against a real PowerPoint install —
  test these first if you touch them.
