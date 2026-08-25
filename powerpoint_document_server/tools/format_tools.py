"""Text and shape formatting tools for live PowerPoint COM automation."""

import json
import sys


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
) -> str:
    """Format the text of a shape (font style, name, size, color).

    Only arguments that are explicitly set are changed.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        shape: Shape name, or its 1-based index on the slide (required).
        bold, italic, underline: Font style toggles.
        font_name: e.g. "Calibri".
        font_size: Point size.
        font_color: Hex RGB, e.g. "FF0000" (with or without leading '#').

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or not shape:
        return json.dumps({"error": "slide_index and shape are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide, find_shape
        from powerpoint_document_server.utils.conversions import hex_to_bgr

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        sl = find_slide(pres, slide_index)
        target = find_shape(sl, shape)

        if not target.HasTextFrame:
            return json.dumps({"error": f"Shape '{shape}' does not have a text frame"})

        font = target.TextFrame.TextRange.Font
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if underline is not None:
            font.Underline = underline
        if font_name is not None:
            font.Name = font_name
        if font_size is not None:
            font.Size = font_size
        if font_color is not None:
            font.Color.RGB = hex_to_bgr(font_color)

        return json.dumps({"success": True, "presentation": pres.Name, "slide_index": slide_index, "shape": target.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})


def format_shape(
    presentation: str = None,
    slide_index: int = None,
    shape: str = None,
    fill_color: str = None,
    line_color: str = None,
    line_weight: float = None,
) -> str:
    """Format a shape's fill and outline.

    Only arguments that are explicitly set are changed.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        shape: Shape name, or its 1-based index on the slide (required).
        fill_color: Hex RGB fill color, e.g. "4472C4" (with or without '#').
        line_color: Hex RGB outline color.
        line_weight: Outline weight in points.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or not shape:
        return json.dumps({"error": "slide_index and shape are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide, find_shape
        from powerpoint_document_server.utils.conversions import hex_to_bgr

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        sl = find_slide(pres, slide_index)
        target = find_shape(sl, shape)

        if fill_color is not None:
            target.Fill.ForeColor.RGB = hex_to_bgr(fill_color)
        if line_color is not None:
            target.Line.ForeColor.RGB = hex_to_bgr(line_color)
        if line_weight is not None:
            target.Line.Weight = line_weight

        return json.dumps({"success": True, "presentation": pres.Name, "slide_index": slide_index, "shape": target.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})
