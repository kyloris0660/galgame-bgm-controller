"""Cached executable icons with deterministic GDI handle cleanup."""
import ctypes
from functools import lru_cache
import os

from PIL import Image, ImageDraw


def placeholder(size=64):
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((4, 4, size - 4, size - 4), radius=12, fill="#4568df")
    draw.ellipse((size * .25, size * .55, size * .48, size * .77), fill="white")
    draw.line((size * .47, size * .66, size * .47, size * .25, size * .73, size * .20),
              fill="white", width=max(2, size // 14))
    return image


@lru_cache(maxsize=256)
def get_exe_icon_pil(exe_path, large=True):
    size = 64 if large else 32
    if not exe_path or not os.path.exists(exe_path):
        return placeholder(size)
    import win32con
    import win32gui
    import win32ui

    icon = ctypes.c_void_p()
    screen = dc = memory = bitmap = old = None
    try:
        result = ctypes.windll.shell32.SHDefExtractIconW(
            ctypes.c_wchar_p(exe_path), 0, 0, ctypes.byref(icon), None, size)
        if result != 0 or not icon.value:
            return placeholder(size)
        screen = win32gui.GetDC(0)
        dc = win32ui.CreateDCFromHandle(screen)
        memory = dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(dc, size, size)
        old = memory.SelectObject(bitmap)
        memory.FillSolidRect((0, 0, size, size), 0)
        win32gui.DrawIconEx(memory.GetSafeHdc(), 0, 0, icon.value, size, size, 0, None, win32con.DI_NORMAL)
        image = Image.frombytes("RGBA", (size, size), bitmap.GetBitmapBits(True), "raw", "BGRA")
        if not image.getchannel("A").getextrema()[1]:
            # Older icons use an AND mask instead of an alpha channel.
            memory.FillSolidRect((0, 0, size, size), 0xFFFFFF)
            win32gui.DrawIconEx(memory.GetSafeHdc(), 0, 0, icon.value, size, size, 0, None, win32con.DI_MASK)
            mask = Image.frombytes("RGB", (size, size), bitmap.GetBitmapBits(True), "raw", "BGRX").convert("L")
            image.putalpha(mask.point(lambda x: 255 - x))
        return image
    except Exception:
        return placeholder(size)
    finally:
        if old and memory:
            memory.SelectObject(old)
        if bitmap:
            win32gui.DeleteObject(bitmap.GetHandle())
        if memory:
            memory.DeleteDC()
        if screen:
            win32gui.ReleaseDC(0, screen)
        if icon.value:
            win32gui.DestroyIcon(icon.value)
