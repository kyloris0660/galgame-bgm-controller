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
    if not exe_path:
        return _placeholder()
    key = (exe_path, large)
    if key in _ICON_CACHE:
        return _ICON_CACHE[key]
    if not HAVE_WIN or not os.path.exists(exe_path):
        img = _placeholder()
        _ICON_CACHE[key] = img
        return img
    img = None
    large_icons = []
    small_icons = []
    try:
        large_icons, small_icons = win32gui.ExtractIconEx(exe_path, 0)
        hicon = (large_icons or small_icons)[0]
        target = 32  # 统一缩放到 32x32 更稳
        ico_x = win32api.GetSystemMetrics(
            win32con.SM_CXICON if large else win32con.SM_CXSMICON
        )
        ico_y = win32api.GetSystemMetrics(
            win32con.SM_CYICON if large else win32con.SM_CYSMICON
        )
        ico_x = max(16, min(256, ico_x))
        ico_y = max(16, min(256, ico_y))
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
        if img.size != (target, target):
            img = img.resize((target, target), Image.LANCZOS)
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
    return Image.new("RGBA", (32, 32), (0, 0, 0, 0))
