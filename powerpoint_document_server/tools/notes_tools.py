"""Speaker notes tools for live PowerPoint COM automation."""

import json
import sys


def get_speaker_notes(presentation: str = None, slide_index: int = None) -> str:
    """Read the speaker notes for a slide.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).

    Returns:
        JSON with the notes text (empty string if there are none).
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

        notes_text = ""
        try:
            notes_shape = slide.NotesPage.Shapes.Placeholders(2)  # standard notes-body placeholder
            if notes_shape.HasTextFrame and notes_shape.TextFrame.HasText:
                notes_text = notes_shape.TextFrame.TextRange.Text
        except Exception:
            pass

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "notes": notes_text,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_speaker_notes(presentation: str = None, slide_index: int = None, text: str = "") -> str:
    """Set the speaker notes for a slide, replacing any existing notes.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        text: New notes text.

    Returns:
        JSON confirmation.
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

        notes_shape = slide.NotesPage.Shapes.Placeholders(2)
        notes_shape.TextFrame.TextRange.Text = text

        return json.dumps({"success": True, "presentation": pres.Name, "slide_index": slide_index})
    except Exception as e:
        return json.dumps({"error": str(e)})
