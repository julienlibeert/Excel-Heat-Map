"""Windows clipboard operations for image export."""

import io

from PIL import Image

# Windows-only imports (lazy loaded)
try:
    import win32clipboard
    import win32con
    HAS_WIN32 = True
except ImportError:
    win32clipboard = None
    win32con = None
    HAS_WIN32 = False


def copy_image_to_clipboard(image_bytes):
    """Copy image data to the Windows clipboard.

    Args:
        image_bytes (bytes): Image data in BMP format (DIB, without header).

    Raises:
        ImportError: If win32clipboard is not available (non-Windows system).
        Exception: If clipboard operations fail.
    """
    if not HAS_WIN32:
        raise ImportError("Clipboard operations require pywin32 (Windows only).")
    
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, image_bytes)
    finally:
        win32clipboard.CloseClipboard()


def figure_to_dib(figure, figure_size=None, export_dpi=300):
    """Convert a Matplotlib figure to Device Independent Bitmap (DIB) format.

    Converts the figure to PNG, then to RGB BMP, and returns the DIB (BMP data
    without the 14-byte file header) suitable for clipboard operations.

    Args:
        figure (Figure): Matplotlib figure to convert.
        figure_size (tuple, optional): (width_inches, height_inches) for export.
            If None, uses the figure's current size.
        export_dpi (int): DPI for export. Default 300 for high quality.

    Returns:
        bytes: DIB data suitable for Windows clipboard.
    """
    png = io.BytesIO()
    previous_size = figure.get_size_inches().copy()
    try:
        if figure_size is not None:
            figure.set_size_inches(*figure_size, forward=False)
        figure.savefig(
            png,
            format="png",
            dpi=export_dpi,
            bbox_inches=None if figure_size is not None else "tight",
        )
    finally:
        if figure_size is not None:
            figure.set_size_inches(*previous_size, forward=False)
    png.seek(0)
    with Image.open(png) as image:
        bitmap = io.BytesIO()
        image.convert("RGB").save(bitmap, format="BMP")
    return bitmap.getvalue()[14:]  # Strip 14-byte BMP file header
