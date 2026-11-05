# powerpoint-mcp-live

**Edit a PowerPoint presentation while it's open — the AI attaches to the window you already
have, it doesn't start its own.**

`Live editing` &middot; `Windows (COM)` &middot; `23 tools`

---

Most PowerPoint MCP servers (including this workspace's own
[mcp-server-powerpoint](../mcp-server-powerpoint)) work by starting a fresh, invisible PowerPoint
process, opening your file in it, and closing it again when done. That's great for scripted deck
generation, but it means the AI is never looking at the same window you are.

`powerpoint-mcp-live` does the opposite: it attaches to whatever PowerPoint window you already have
open via COM, using the same "find the running application" approach as
[word-mcp-live](../word-mcp-live) and [excel-mcp-live](../excel-mcp-live). It never launches
PowerPoint and never closes it — it only touches presentations you opened yourself.

## Requirements

- **Windows** (COM automation is Windows-only — there is no macOS/Linux mode for this project)
- **Microsoft PowerPoint** installed and **already running with a presentation open**
- **Python 3.11+**

## Installation

```bash
git clone <this-repo-url>
cd powerpoint-mcp-live
pip install -e .
```

Or run straight from source with [uv](https://docs.astral.sh/uv/) — no install step needed:

```bash
uv run powerpoint_mcp_server.py
```

## Client setup (`.mcp.json`)

Point your MCP client at the project directory. For Claude Code, add this to `.mcp.json` in your
project (or `~/.claude.json` for a user-wide config):

```json
{
  "mcpServers": {
    "powerpoint-live": {
      "command": "uv",
      "args": [
        "--directory",
        "G:/Ai/MCP/powerpoint-mcp-live",
        "run",
        "powerpoint_mcp_server.py"
      ],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

If you installed it with `pip install -e .` instead, you can use the console script directly:

```json
{
  "mcpServers": {
    "powerpoint-live": {
      "command": "powerpoint-mcp-live",
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

`MCP_AUTHOR` / `MCP_AUTHOR_INITIALS` set the author name/initials attached to new slide comments
(default: `"Author"` / `""`).

### Limit tool discovery

Set `POWERPOINT_MCP_TOOLSETS` in the server's `env` to a comma-separated list of groups to register
only those tools, e.g. `"POWERPOINT_MCP_TOOLSETS": "slides,shapes"`. Groups are combined as a
union. Available groups are `presentations`, `slides`, `shapes` (shapes and text), `formatting`,
`notes`, `export` and `comments`, matching the sections of the table below. Each tool is in exactly
one group. `list_open_presentations` is always registered. Unset, empty or `all` registers every
tool. An unknown group name stops the server at startup. Restart the server after changing this
setting. Filtering only controls which tools are listed and callable. It is not a security
boundary.

## Supported functions

Every tool operates on a presentation that's already open in PowerPoint. `presentation` is
optional almost everywhere — omit it to target the active presentation.

| Tool | Description |
|---|---|
| **Presentation** | |
| `list_open_presentations` | List every presentation open in PowerPoint |
| `get_presentation_info` | Slide count, per-slide titles, active slide |
| `save_presentation` | Save in place (Ctrl+S) |
| **Slides** | |
| `list_slides` | List slide indexes and titles |
| `add_slide` | Add a slide with a chosen layout |
| `delete_slide` | Delete a slide |
| `duplicate_slide` | Duplicate a slide, inserted right after the original |
| `move_slide` | Reorder a slide to a new position |
| `activate_slide` | Make a slide the visible one in the PowerPoint window |
| **Shapes & text** | |
| `list_shapes` | List shapes on a slide (name, type, position, text) |
| `get_slide_text` | Read a slide's title and all body text |
| `add_textbox` | Add a text box at a given position |
| `set_shape_text` | Set the text of an existing shape/placeholder |
| `delete_shape` | Delete a shape |
| `set_shape_position` | Move and/or resize a shape |
| **Formatting** | |
| `format_shape_text` | Bold/italic/underline, font, size, color |
| `format_shape` | Fill color, outline color, outline weight |
| **Speaker notes** | |
| `get_speaker_notes` | Read a slide's speaker notes |
| `set_speaker_notes` | Replace a slide's speaker notes |
| **Export** | |
| `export_slide_to_image` | Export a slide to PNG/JPG/BMP/GIF for visual verification |
| **Comments** | |
| `add_comment` | Add a slide comment (PowerPoint 2016+) |
| `get_comments` | List comments on a slide |
| `delete_comment` | Remove a comment |

Not yet covered: tables, charts, SmartArt, animations, transitions, templates/themes — for those,
use [mcp-server-powerpoint](../mcp-server-powerpoint), which starts its own PowerPoint instance and
has much broader operation coverage.

## Example prompts

```
"I have this deck open — read the text on slide 3 and tell me if the bullets are too long."
"Add a new slide after slide 2 with a blank layout."
"Add a text box on slide 4 that says 'Draft — do not distribute' in red."
"Make the title on slide 1 bold and increase it to 44pt."
"Export slide 5 as an image so I can check whether anything overlaps."
"Add a comment on slide 3 asking if the chart numbers are up to date."
```

## Known limitations

- **Windows only.** PowerPoint COM automation doesn't exist on macOS/Linux.
- **The presentation must already be open in PowerPoint.** This server doesn't create or open
  files — see mcp-server-powerpoint for that.
- **No single-undo grouping.** Word's COM API can wrap several edits into one Ctrl+Z; PowerPoint's
  cannot (neither can Excel's). A tool call that performs multiple writes may need several Ctrl+Z's
  to fully undo.
- **Comments require PowerPoint 2016+** (modern threaded comments).

## Development

See [CLAUDE.md](CLAUDE.md) for architecture notes and design constraints — in particular, why this
server must never launch or quit PowerPoint itself.
