"""Slide management tools for live PowerPoint COM automation."""

import json
import sys

# Classic PpSlideLayout constants (pre-2007 enum, still accepted by Slides.Add)
_LAYOUTS = {
    "title": 1,        # ppLayoutTitle
    "text": 2,         # ppLayoutText
    "two_content": 4,  # ppLayoutTwoColumnText
    "title_only": 11,  # ppLayoutTitleOnly
    "blank": 12,        # ppLayoutBlank
}


def list_slides(presentation: str = None) -> str:
    """List slides in an open presentation, with index and title text.

    Args:
        presentation: Presentation name or path (None = active presentation).

    Returns:
        JSON with each slide's index and title (if it has one).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)

        slides = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            title = slide.Shapes.Title.TextFrame.TextRange.Text if slide.Shapes.HasTitle else None
            slides.append({"index": i, "title": title})

        return json.dumps({"success": True, "presentation": pres.Name, "slides": slides}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def add_slide(presentation: str = None, index: int = None, layout: str = "text") -> str:
    """Add a new slide to an open presentation.

    Args:
        presentation: Presentation name or path (None = active presentation).
        index: 1-based position to insert at (None = append at the end).
        layout: One of "title", "text", "two_content", "title_only", "blank".

    Returns:
        JSON with the new slide's index.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    layout_key = layout or "text"
    if layout_key not in _LAYOUTS:
        return json.dumps({"error": f"Invalid layout: {layout}. Use one of {list(_LAYOUTS)}"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        target_index = index if index is not None else pres.Slides.Count + 1

        new_slide = pres.Slides.Add(target_index, _LAYOUTS[layout_key])

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": new_slide.SlideIndex,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_slide(presentation: str = None, index: int = None) -> str:
    """Delete a slide from an open presentation.

    Args:
        presentation: Presentation name or path (None = active presentation).
        index: 1-based slide index (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if index is None:
        return json.dumps({"error": "index is required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, index)
        slide.Delete()

        return json.dumps({"success": True, "presentation": pres.Name, "deleted_index": index})
    except Exception as e:
        return json.dumps({"error": str(e)})


def duplicate_slide(presentation: str = None, index: int = None) -> str:
    """Duplicate a slide, inserting the copy immediately after the original.

    Args:
        presentation: Presentation name or path (None = active presentation).
        index: 1-based slide index to duplicate (required).

    Returns:
        JSON with the new slide's index.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if index is None:
        return json.dumps({"error": "index is required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, index)
        new_range = slide.Duplicate()  # returns a SlideRange containing the copy
        new_slide = new_range(1)

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "source_index": index,
            "new_index": new_slide.SlideIndex,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def move_slide(presentation: str = None, index: int = None, new_index: int = None) -> str:
    """Move a slide to a new position in the slide order.

    Args:
        presentation: Presentation name or path (None = active presentation).
        index: Current 1-based slide index (required).
        new_index: Target 1-based position (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if index is None or new_index is None:
        return json.dumps({"error": "index and new_index are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, index)
        slide.MoveTo(new_index)

        return json.dumps({"success": True, "presentation": pres.Name, "moved_to": new_index})
    except Exception as e:
        return json.dumps({"error": str(e)})


def activate_slide(presentation: str = None, index: int = None) -> str:
    """Make a slide the visible slide in PowerPoint's active window.

    Args:
        presentation: Presentation name or path (None = active presentation).
        index: 1-based slide index (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if index is None:
        return json.dumps({"error": "index is required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, index)
        app.ActiveWindow.View.GotoSlide(slide.SlideIndex)

        return json.dumps({"success": True, "presentation": pres.Name, "active_slide_index": index})
    except Exception as e:
        return json.dumps({"error": str(e)})
