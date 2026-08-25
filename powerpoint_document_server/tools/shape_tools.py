"""Shape and text tools for live PowerPoint COM automation."""

import json
import sys


def list_shapes(presentation: str = None, slide_index: int = None) -> str:
    """List shapes on a slide with name, type, position, and text (if any).

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).

    Returns:
        JSON with each shape's 1-based index, name, MsoShapeType, position/
        size (points), and text.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None:
        return json.dumps({"error": "slide_index is required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, slide_index)

        shapes = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            text = None
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text
            shapes.append({
                "index": i,
                "name": shape.Name,
                "type": shape.Type,
                "left": shape.Left,
                "top": shape.Top,
                "width": shape.Width,
                "height": shape.Height,
                "text": text,
            })

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "shapes": shapes,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_slide_text(presentation: str = None, slide_index: int = None) -> str:
    """Read all text on a slide — the title plus every other text-bearing shape.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).

    Returns:
        JSON with the title and a list of body text strings, in shape order.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None:
        return json.dumps({"error": "slide_index is required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, slide_index)

        has_title = slide.Shapes.HasTitle
        title = slide.Shapes.Title.TextFrame.TextRange.Text if has_title else None
        title_id = slide.Shapes.Title.Id if has_title else None

        body = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if title_id is not None and shape.Id == title_id:
                continue
            if shape.HasTextFrame and shape.TextFrame.HasText:
                body.append(shape.TextFrame.TextRange.Text)

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "title": title,
            "body_text": body,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def add_textbox(
    presentation: str = None,
    slide_index: int = None,
    text: str = "",
    left: float = 50,
    top: float = 50,
    width: float = 300,
    height: float = 50,
) -> str:
    """Add a text box to a slide.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        text: Text content.
        left, top, width, height: Position and size in points.

    Returns:
        JSON with the new shape's name and 1-based index on the slide.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None:
        return json.dumps({"error": "slide_index is required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, slide_index)

        MSO_TEXT_ORIENTATION_HORIZONTAL = 1
        shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)
        if text:
            shape.TextFrame.TextRange.Text = text

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "shape_name": shape.Name,
            "shape_index": slide.Shapes.Count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_shape_text(presentation: str = None, slide_index: int = None, shape: str = None, text: str = "") -> str:
    """Set the text of an existing shape (title, placeholder, or text box).

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        shape: Shape name, or its 1-based index on the slide (required).
        text: New text content.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or not shape:
        return json.dumps({"error": "slide_index and shape are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide, find_shape

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        sl = find_slide(pres, slide_index)
        target = find_shape(sl, shape)

        if not target.HasTextFrame:
            return json.dumps({"error": f"Shape '{shape}' does not have a text frame"})
        target.TextFrame.TextRange.Text = text

        return json.dumps({"success": True, "presentation": pres.Name, "slide_index": slide_index, "shape": target.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_shape(presentation: str = None, slide_index: int = None, shape: str = None) -> str:
    """Delete a shape from a slide.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        shape: Shape name, or its 1-based index on the slide (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or not shape:
        return json.dumps({"error": "slide_index and shape are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide, find_shape

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        sl = find_slide(pres, slide_index)
        target = find_shape(sl, shape)
        name = target.Name
        target.Delete()

        return json.dumps({"success": True, "presentation": pres.Name, "slide_index": slide_index, "deleted": name})
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_shape_position(
    presentation: str = None,
    slide_index: int = None,
    shape: str = None,
    left: float = None,
    top: float = None,
    width: float = None,
    height: float = None,
) -> str:
    """Move and/or resize a shape. Only provided arguments are changed.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        shape: Shape name, or its 1-based index on the slide (required).
        left, top, width, height: New position/size in points (None = unchanged).

    Returns:
        JSON with the shape's resulting position and size.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or not shape:
        return json.dumps({"error": "slide_index and shape are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide, find_shape

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        sl = find_slide(pres, slide_index)
        target = find_shape(sl, shape)

        if left is not None:
            target.Left = left
        if top is not None:
            target.Top = top
        if width is not None:
            target.Width = width
        if height is not None:
            target.Height = height

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "shape": target.Name,
            "left": target.Left,
            "top": target.Top,
            "width": target.Width,
            "height": target.Height,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
