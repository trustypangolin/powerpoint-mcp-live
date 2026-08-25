"""Slide comment tools for live PowerPoint COM automation (PowerPoint 2016+)."""

import json
import sys

from powerpoint_document_server.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS


def add_comment(
    presentation: str = None,
    slide_index: int = None,
    text: str = "",
    left: float = 50,
    top: float = 50,
    author: str = None,
    author_initials: str = None,
) -> str:
    """Add a comment to a slide (PowerPoint 2016+ modern comments).

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        text: Comment text (required).
        left, top: Position of the comment marker, in points.
        author: Comment author (default: MCP_AUTHOR env var, or "Author").
        author_initials: Author initials (default: MCP_AUTHOR_INITIALS env var).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or not text:
        return json.dumps({"error": "slide_index and text are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, slide_index)

        author_name = author or DEFAULT_AUTHOR
        initials = author_initials if author_initials is not None else DEFAULT_INITIALS

        slide.Comments.Add(left, top, author_name, initials, text)

        return json.dumps({"success": True, "presentation": pres.Name, "slide_index": slide_index})
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_comments(presentation: str = None, slide_index: int = None) -> str:
    """List all comments on a slide.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).

    Returns:
        JSON with each comment's 1-based index, author, text, and position.
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

        comments = []
        for i in range(1, slide.Comments.Count + 1):
            com = slide.Comments(i)
            comments.append({
                "index": i,
                "author": com.Author,
                "author_initials": com.AuthorInitials,
                "text": com.Text,
                "left": com.Left,
                "top": com.Top,
            })

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "comments": comments,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_comment(presentation: str = None, slide_index: int = None, index: int = None) -> str:
    """Delete a comment from a slide.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        index: 1-based comment index, from get_comments (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or index is None:
        return json.dumps({"error": "slide_index and index are required"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, slide_index)

        if index < 1 or index > slide.Comments.Count:
            return json.dumps({"error": f"Comment index {index} out of range. Slide has {slide.Comments.Count} comment(s)."})

        slide.Comments(index).Delete()

        return json.dumps({"success": True, "presentation": pres.Name, "slide_index": slide_index, "deleted_index": index})
    except Exception as e:
        return json.dumps({"error": str(e)})
