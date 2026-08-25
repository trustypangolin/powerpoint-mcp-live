"""COM connection manager for Microsoft PowerPoint on Windows.

Provides functions to attach to a running PowerPoint instance and locate
open presentations/slides/shapes. Only works on Windows with pywin32
installed.

Unlike mcp-server-powerpoint (the sibling project), this module never
launches or quits a PowerPoint.Application — it only attaches to a process
the user already started. Every lookup here must remain read-only with
respect to process lifetime: no Quit(), no Close() on a presentation the
caller didn't explicitly ask to close.

Like Word (and unlike Excel), PowerPoint is effectively single-instance —
opening multiple presentations opens multiple windows within one
PowerPoint.Application process, not multiple separate processes. So this
module mirrors word-mcp-live's core/word_com.py single-app-instance pattern
rather than excel-mcp-live's multi-instance Running Object Table walk.
"""

import os
import sys
import unicodedata


def get_powerpoint_app():
    """Get a reference to the running PowerPoint application via COM.

    Returns the PowerPoint.Application COM object that has open
    presentations. Raises RuntimeError if PowerPoint is not running or not
    on Windows.
    """
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint COM automation is only available on Windows")

    import pythoncom
    import win32com.client

    # COM is apartment-threaded: each OS thread must initialize its own
    # apartment before touching COM objects. FastMCP may dispatch this call
    # on a fresh worker thread per request, so this can't be a one-time
    # process-level init — it must run on every call. CoInitialize() is a
    # no-op (returns S_FALSE) if the thread is already initialized, so this
    # is safe to call unconditionally.
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
        if app.Presentations.Count > 0:
            return app
        app_with_decks = _find_powerpoint_with_presentations()
        if app_with_decks is not None:
            return app_with_decks
        return app
    except Exception:
        app_with_decks = _find_powerpoint_with_presentations()
        if app_with_decks is not None:
            return app_with_decks
        raise RuntimeError(
            "Microsoft PowerPoint is not running. Please open PowerPoint first."
        )


def _find_powerpoint_with_presentations():
    """Scan the Running Object Table for a PowerPoint.Application with open decks.

    Handles cases where GetActiveObject returns an empty Application proxy.
    Presentations are also registered in the ROT directly as file monikers
    (.pptx/.ppt/.pptm paths); in that case we grab the Presentation COM
    object from the moniker and reach the owning Application via
    ``presentation.Application``.

    Returns the PowerPoint.Application COM object if found, or None.
    """
    try:
        import pythoncom
        import win32com.client

        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()

        monikers_to_retry = []
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                if hasattr(com_obj, "Presentations") and hasattr(com_obj, "ActivePresentation"):
                    if com_obj.Presentations.Count > 0:
                        return com_obj
                if name and name.lower().endswith((".pptx", ".ppt", ".pptm")):
                    monikers_to_retry.append((name, moniker))
            except Exception:
                try:
                    ctx = pythoncom.CreateBindCtx(0)
                    name = moniker.GetDisplayName(ctx, None)
                    if name and name.lower().endswith((".pptx", ".ppt", ".pptm")):
                        monikers_to_retry.append((name, moniker))
                except Exception:
                    pass
                continue

        for name, moniker in monikers_to_retry:
            try:
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                pres = win32com.client.Dispatch(dispatch)
                app = pres.Application
                if app.Presentations.Count > 0:
                    return app
            except Exception:
                continue
    except Exception:
        pass
    return None


def find_presentation(app, name: str = None):
    """Find an open presentation by filename.

    Args:
        app: PowerPoint.Application COM object.
        name: Presentation name (basename) or full path. If None or empty,
              returns the active presentation.

    Returns:
        Presentation COM object.

    Raises:
        ValueError: If the presentation is not found or none are open.
    """
    if app.Presentations.Count == 0:
        raise ValueError("No presentations are open in PowerPoint")

    if not name:
        return app.ActivePresentation

    target_basename = unicodedata.normalize("NFC", os.path.basename(name)).lower()
    target_fullpath = (
        unicodedata.normalize("NFC", os.path.normpath(name)).lower()
        if os.path.isabs(name)
        else None
    )

    for i in range(1, app.Presentations.Count + 1):
        pres = app.Presentations(i)
        if unicodedata.normalize("NFC", pres.Name).lower() == target_basename:
            return pres
        if target_fullpath and unicodedata.normalize("NFC", os.path.normpath(pres.FullName)).lower() == target_fullpath:
            return pres

    open_decks = [app.Presentations(i).Name for i in range(1, app.Presentations.Count + 1)]
    raise ValueError(
        f"Presentation '{name}' is not open in PowerPoint. Open presentations: {open_decks}"
    )


def find_slide(presentation, index: int):
    """Find a slide by its 1-based index in an open presentation.

    Args:
        presentation: Presentation COM object.
        index: 1-based slide index.

    Returns:
        Slide COM object.

    Raises:
        ValueError: If the index is out of range.
    """
    count = presentation.Slides.Count
    if index is None or index < 1 or index > count:
        raise ValueError(f"Slide index {index} out of range. Presentation has {count} slide(s).")
    return presentation.Slides(index)


def find_shape(slide, shape_ref):
    """Find a shape on a slide by 1-based index or by name.

    Args:
        slide: Slide COM object.
        shape_ref: Shape name (str), or a 1-based index (int, or numeric str).

    Returns:
        Shape COM object.

    Raises:
        ValueError: If no shape matches.
    """
    if isinstance(shape_ref, int) or (isinstance(shape_ref, str) and shape_ref.isdigit()):
        index = int(shape_ref)
        if index < 1 or index > slide.Shapes.Count:
            raise ValueError(f"Shape index {index} out of range. Slide has {slide.Shapes.Count} shape(s).")
        return slide.Shapes(index)

    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.Name == shape_ref:
            return shape

    names = [slide.Shapes(i).Name for i in range(1, slide.Shapes.Count + 1)]
    raise ValueError(f"Shape '{shape_ref}' not found on slide. Available shapes: {names}")
