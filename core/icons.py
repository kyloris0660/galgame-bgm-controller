# core/icons.py
from PIL import Image
import os

try:
    import win32gui
    import win32api
    import win32con
    import win32ui

    HAVE_WIN = True
except Exception:
    HAVE_WIN = False

_ICON_CACHE = {}


def get_exe_icon_pil(exe_path: str, large=True) -> Image.Image:
    """从 exe 抽取图标（PIL），带缓存；失败返回占位图。"""
    if not exe_path:
        return _placeholder()
    key = (exe_path, large)
    if key in _ICON_CACHE:
        return _ICON_CACHE[key]
    if not HAVE_WIN or not os.path.exists(exe_path):
        img = _placeholder()
        _ICON_CACHE[key] = img
        return img
    try:
        large_icons, small_icons = win32gui.ExtractIconEx(exe_path, 0)
        hicon = (large_icons or small_icons)[0]
        ico_x = win32api.GetSystemMetrics(
            win32con.SM_CXICON if large else win32con.SM_CXSMICON
        )
        ico_y = win32api.GetSystemMetrics(
            win32con.SM_CYICON if large else win32con.SM_CYSMICON
        )
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
        hcdc = hdc.CreateCompatibleDC()
        hcdc.SelectObject(hbmp)
        win32gui.DrawIconEx(
            hcdc.GetSafeHdc(), 0, 0, hicon, ico_x, ico_y, 0, None, win32con.DI_NORMAL
        )
        bmpinfo = hbmp.GetInfo()
        bmpstr = hbmp.GetBitmapBits(True)
        img = Image.frombuffer(
            "RGB",
            (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
            bmpstr,
            "raw",
            "BGRX",
            0,
            1,
        ).convert("RGBA")
    except Exception:
        img = _placeholder()
    finally:
        try:
            for hi in large_icons + small_icons:
                win32gui.DestroyIcon(hi)
        except Exception:
            pass
    _ICON_CACHE[key] = img
    return img


def _placeholder() -> Image.Image:
    return Image.new("RGBA", (48, 48), (0, 0, 0, 0))
