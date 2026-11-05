"""
Main entry point for the PowerPoint Live MCP Server.

Live-attach tools that operate on presentations already open in PowerPoint
via COM automation. Requires Windows with PowerPoint installed and running —
this server never launches or quits PowerPoint itself.
"""

import os
import sys

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from powerpoint_document_server.tool_metadata import SERVER_INSTRUCTIONS
from powerpoint_document_server.toolsets import ENV_VAR, ToolRegistrar, parse_toolsets
from powerpoint_document_server.tools import (
    presentation_tools,
    slide_tools,
    shape_tools,
    format_tools,
    notes_tools,
    export_tools,
    comment_tools,
)

mcp = FastMCP("PowerPoint Live MCP Server", instructions=SERVER_INSTRUCTIONS)


def register_tools(server=None, toolsets=None):
    """Register tools with the MCP server.

    ``description=impl.__doc__`` passes each tool module's docstring through
    ``ToolRegistrar``, which turns its ``Args:`` section into per-parameter
    schema descriptions and keeps only the summary as the tool description.
    ``toolsets`` (default: the POWERPOINT_MCP_TOOLSETS env var) limits which
    groups are registered.
    """
    if server is None:
        server = globals()["mcp"]
    if toolsets is None:
        toolsets = parse_toolsets(os.getenv(ENV_VAR))
    mcp = ToolRegistrar(server, toolsets)

    # --- Presentation tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="List Open Presentations", readOnlyHint=True),
        description=presentation_tools.list_open_presentations.__doc__,
    )
    def list_open_presentations():
        return presentation_tools.list_open_presentations()

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Presentation Info", readOnlyHint=True),
        description=presentation_tools.get_presentation_info.__doc__,
    )
    def get_presentation_info(presentation: str = None):
        return presentation_tools.get_presentation_info(presentation)

    @mcp.tool(
        annotations=ToolAnnotations(title="Save Presentation"),
        description=presentation_tools.save_presentation.__doc__,
    )
    def save_presentation(presentation: str = None):
        return presentation_tools.save_presentation(presentation)

    # --- Slide tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="List Slides", readOnlyHint=True),
        description=slide_tools.list_slides.__doc__,
    )
    def list_slides(presentation: str = None):
        return slide_tools.list_slides(presentation)

    @mcp.tool(
        annotations=ToolAnnotations(title="Add Slide"),
        description=slide_tools.add_slide.__doc__,
    )
    def add_slide(presentation: str = None, index: int = None, layout: str = "text"):
        return slide_tools.add_slide(presentation, index, layout)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Slide", destructiveHint=True),
        description=slide_tools.delete_slide.__doc__,
    )
    def delete_slide(presentation: str = None, index: int = None):
        return slide_tools.delete_slide(presentation, index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Duplicate Slide"),
        description=slide_tools.duplicate_slide.__doc__,
    )
    def duplicate_slide(presentation: str = None, index: int = None):
        return slide_tools.duplicate_slide(presentation, index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Move Slide"),
        description=slide_tools.move_slide.__doc__,
    )
    def move_slide(presentation: str = None, index: int = None, new_index: int = None):
        return slide_tools.move_slide(presentation, index, new_index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Activate Slide"),
        description=slide_tools.activate_slide.__doc__,
    )
    def activate_slide(presentation: str = None, index: int = None):
        return slide_tools.activate_slide(presentation, index)

    # --- Shape / text tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="List Shapes", readOnlyHint=True),
        description=shape_tools.list_shapes.__doc__,
    )
    def list_shapes(presentation: str = None, slide_index: int = None):
        return shape_tools.list_shapes(presentation, slide_index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Slide Text", readOnlyHint=True),
        description=shape_tools.get_slide_text.__doc__,
    )
    def get_slide_text(presentation: str = None, slide_index: int = None):
        return shape_tools.get_slide_text(presentation, slide_index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Add Text Box"),
        description=shape_tools.add_textbox.__doc__,
    )
    def add_textbox(
        presentation: str = None,
        slide_index: int = None,
        text: str = "",
        left: float = 50,
        top: float = 50,
        width: float = 300,
        height: float = 50,
    ):
        return shape_tools.add_textbox(presentation, slide_index, text, left, top, width, height)

    @mcp.tool(
        annotations=ToolAnnotations(title="Set Shape Text"),
        description=shape_tools.set_shape_text.__doc__,
    )
    def set_shape_text(presentation: str = None, slide_index: int = None, shape: str = None, text: str = ""):
        return shape_tools.set_shape_text(presentation, slide_index, shape, text)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Shape", destructiveHint=True),
        description=shape_tools.delete_shape.__doc__,
    )
    def delete_shape(presentation: str = None, slide_index: int = None, shape: str = None):
        return shape_tools.delete_shape(presentation, slide_index, shape)

    @mcp.tool(
        annotations=ToolAnnotations(title="Set Shape Position"),
        description=shape_tools.set_shape_position.__doc__,
    )
    def set_shape_position(
        presentation: str = None,
        slide_index: int = None,
        shape: str = None,
        left: float = None,
        top: float = None,
        width: float = None,
        height: float = None,
    ):
        return shape_tools.set_shape_position(presentation, slide_index, shape, left, top, width, height)

    # --- Formatting tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Format Shape Text"),
        description=format_tools.format_shape_text.__doc__,
    )
    def format_shape_text(
        presentation: str = None,
        slide_index: int = None,
        shape: str = None,
        bold: bool = None,
        italic: bool = None,
        underline: bool = None,
        font_name: str = None,
        font_size: float = None,
        font_color: str = None,
    ):
        return format_tools.format_shape_text(
            presentation, slide_index, shape, bold, italic, underline, font_name, font_size, font_color,
        )

    @mcp.tool(
        annotations=ToolAnnotations(title="Format Shape"),
        description=format_tools.format_shape.__doc__,
    )
    def format_shape(
        presentation: str = None,
        slide_index: int = None,
        shape: str = None,
        fill_color: str = None,
        line_color: str = None,
        line_weight: float = None,
    ):
        return format_tools.format_shape(presentation, slide_index, shape, fill_color, line_color, line_weight)

    # --- Speaker notes tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Speaker Notes", readOnlyHint=True),
        description=notes_tools.get_speaker_notes.__doc__,
    )
    def get_speaker_notes(presentation: str = None, slide_index: int = None):
        return notes_tools.get_speaker_notes(presentation, slide_index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Set Speaker Notes"),
        description=notes_tools.set_speaker_notes.__doc__,
    )
    def set_speaker_notes(presentation: str = None, slide_index: int = None, text: str = ""):
        return notes_tools.set_speaker_notes(presentation, slide_index, text)

    # --- Export tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Export Slide to Image", readOnlyHint=True),
        description=export_tools.export_slide_to_image.__doc__,
    )
    def export_slide_to_image(
        presentation: str = None,
        slide_index: int = None,
        output_path: str = None,
        width: int = 1920,
        height: int = 1080,
    ):
        return export_tools.export_slide_to_image(presentation, slide_index, output_path, width, height)

    # --- Comment tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Add Comment"),
        description=comment_tools.add_comment.__doc__,
    )
    def add_comment(
        presentation: str = None,
        slide_index: int = None,
        text: str = "",
        left: float = 50,
        top: float = 50,
        author: str = None,
        author_initials: str = None,
    ):
        return comment_tools.add_comment(presentation, slide_index, text, left, top, author, author_initials)

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Comments", readOnlyHint=True),
        description=comment_tools.get_comments.__doc__,
    )
    def get_comments(presentation: str = None, slide_index: int = None):
        return comment_tools.get_comments(presentation, slide_index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Comment", destructiveHint=True),
        description=comment_tools.delete_comment.__doc__,
    )
    def delete_comment(presentation: str = None, slide_index: int = None, index: int = None):
        return comment_tools.delete_comment(presentation, slide_index, index)


def run_server():
    """Run the PowerPoint Live MCP Server over stdio."""
    register_tools()
    print("PowerPoint Live MCP Server starting (stdio transport)...", file=sys.stderr)
    mcp.run(transport="stdio")


def main():
    """Main entry point for the server."""
    run_server()


if __name__ == "__main__":
    main()
