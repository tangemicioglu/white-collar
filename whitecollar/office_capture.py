from __future__ import annotations

import ctypes
import time
from pathlib import Path

from .errors import BackendUnavailableError


def capture_window(hwnd: int, output: Path) -> None:
    """Capture a native Office window, including hardware-rendered content."""

    try:
        from PIL import Image, ImageGrab
    except ImportError as exc:
        raise BackendUnavailableError("office-screen-capture") from exc

    try:
        import win32gui

        win32gui.ShowWindow(hwnd, 9)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.2)
    image = None
    try:
        import win32gui
        import win32ui

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width, height = right - left, bottom - top
        if width <= 0 or height <= 0:
            raise RuntimeError("Office window has no visible dimensions")
        window_dc = win32gui.GetWindowDC(hwnd)
        source_dc = win32ui.CreateDCFromHandle(window_dc)
        memory_dc = source_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(source_dc, width, height)
        memory_dc.SelectObject(bitmap)
        try:
            rendered = ctypes.windll.user32.PrintWindow(hwnd, memory_dc.GetSafeHdc(), 2)
            if not rendered:
                raise RuntimeError("PrintWindow could not render the Office window")
            bits = bitmap.GetBitmapBits(True)
            image = Image.frombuffer("RGB", (width, height), bits, "raw", "BGRX", 0, 1)
        finally:
            win32gui.DeleteObject(bitmap.GetHandle())
            memory_dc.DeleteDC()
            source_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, window_dc)
    except Exception:
        image = ImageGrab.grab(window=hwnd)
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output, format="PNG")
