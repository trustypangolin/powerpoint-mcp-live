"""Presentation-level tools for live PowerPoint COM automation."""

import json
import sys


def list_open_presentations() -> str:
    """List all presentations currently open in PowerPoint.

    Returns:
        JSON with each presentation's name, full path, save state, slide
        count, and whether it's the active presentation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app

        app = get_powerpoint_app()
        active_name = app.ActivePresentation.Name if app.Presentations.Count > 0 else None

        presentations = []
        for i in range(1, app.Presentations.Count + 1):
            pres = app.Presentations(i)
            presentations.append({
                "name": pres.Name,
                "full_path": pres.FullName,
                "saved": bool(pres.Saved),
                "slide_count": pres.Slides.Count,
                "is_active": pres.Name == active_name,
            })

        return json.dumps({"success": True, "presentations": presentations}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_presentation_info(presentation: str = None) -> str:
    """Get slide count, per-slide titles, and the active slide for an open presentation.

    Args:
        presentation: Presentation name or path (None = active presentation).

    Returns:
        JSON with each slide's index, title text (if it has a title
        placeholder), and layout.
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
            slides.append({"index": i, "title": title, "layout": slide.Layout})

        active_index = None
        try:
            active_index = app.ActiveWindow.View.Slide.SlideIndex
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "full_path": pres.FullName,
            "slide_count": pres.Slides.Count,
            "active_slide_index": active_index,
            "slides": slides,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def save_presentation(presentation: str = None) -> str:
    """Save an open presentation in place (equivalent to pressing Ctrl+S).

    Args:
        presentation: Presentation name or path (None = active presentation).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        pres.Save()

        return json.dumps({"success": True, "presentation": pres.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})
