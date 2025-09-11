# ui/tray.py
from PIL import Image, ImageDraw
from core.icons import get_exe_icon_pil

try:
    import pystray

    HAVE = True
except Exception:
    HAVE = False


def _default_img():
    img = Image.new("RGBA", (64, 64), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((8, 8, 56, 56), outline=(0, 0, 0), width=3)
    d.rectangle((24, 24, 40, 40), fill=(0, 0, 0))
    return img


class Tray:
    def __init__(self, on_select, on_pause, on_quit, on_mode_change, get_mode):
        self.on_select = on_select
        self.on_pause = on_pause
        self.on_quit = on_quit
        self.on_mode_change = on_mode_change
        self.get_mode = get_mode
        self.icon = None
        self._img_default = _default_img()

    def start(self):
        if not HAVE or self.icon:
            return

        def _mode_label(item):
            return f"静音规则（当前：{self.get_mode()}）"

        def _to_not_fg(icon, item):
            self.on_mode_change("not_foreground")
            self.icon.update_menu()

        def _to_min(icon, item):
            self.on_mode_change("minimized_only")
            self.icon.update_menu()

        menu = pystray.Menu(
            pystray.MenuItem("选择监听软件", lambda: self.on_select() or None),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                _mode_label,
                pystray.Menu(
                    pystray.MenuItem("不在前台时静音", _to_not_fg),
                    pystray.MenuItem("仅最小化时静音", _to_min),
                ),
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("暂停/继续", lambda: self.on_pause() or None),
            pystray.MenuItem("退出", lambda: self.on_quit() or None),
        )
        self.icon = pystray.Icon(
            "bgm_controller", self._img_default, "BGM Controller", menu
        )
        self.icon.run_detached()

    def stop(self):
        if self.icon:
            try:
                self.icon.stop()
            except Exception:
                pass

    def set_icon_from_exe(self, exe_path: str | None):
        """将托盘图标设为该 exe 的图标；传 None/空则回到默认。"""
        if not self.icon:
            return
        if not exe_path:
            self.icon.icon = self._img_default
            self.icon.visible = True
            return
        pil = get_exe_icon_pil(exe_path, large=True)
        self.icon.icon = pil
        self.icon.visible = True
