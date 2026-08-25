"""Small value conversions for PowerPoint COM automation."""


def hex_to_bgr(hex_color: str) -> int:
    """Convert a '#RRGGBB' or 'RRGGBB' hex string to the BGR integer that
    PowerPoint's ColorFormat.RGB property expects (legacy OLE COLORREF byte
    order, not RGB).
    """
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid hex color: {hex_color!r}. Expected 'RRGGBB' or '#RRGGBB'.")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r
