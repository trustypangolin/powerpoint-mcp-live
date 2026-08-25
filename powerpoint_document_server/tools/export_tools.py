"""Slide export tools for live PowerPoint COM automation."""

import json
import os
import sys


def export_slide_to_image(
    presentation: str = None,
    slide_index: int = None,
    output_path: str = None,
    width: int = 1920,
    height: int = 1080,
) -> str:
    """Export a slide to an image file for visual verification.

    Args:
        presentation: Presentation name or path (None = active presentation).
        slide_index: 1-based slide index (required).
        output_path: Destination file path (required). Extension selects
                     the format: .png, .jpg, .bmp, or .gif.
        width, height: Output image dimensions in pixels.

    Returns:
        JSON with the exported file's absolute path. A vision-capable
        assistant can read the image back to check for overlapping shapes,
        text overflow, or layout regressions after an edit.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if slide_index is None or not output_path:
        return json.dumps({"error": "slide_index and output_path are required"})

    ext = os.path.splitext(output_path)[1].lstrip(".").upper()
    if ext not in ("PNG", "JPG", "BMP", "GIF"):
        return json.dumps({"error": f"Unsupported image extension '.{ext}'. Use .png, .jpg, .bmp, or .gif"})

    try:
        from powerpoint_document_server.core.powerpoint_com import get_powerpoint_app, find_presentation, find_slide

        app = get_powerpoint_app()
        pres = find_presentation(app, presentation)
        slide = find_slide(pres, slide_index)

        full_output = os.path.abspath(output_path)
        directory = os.path.dirname(full_output)
        if directory and not os.path.isdir(directory):
            return json.dumps({"error": f"Output directory does not exist: {directory}"})

        slide.Export(full_output, ext, width, height)

        return json.dumps({
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "output_path": full_output,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
