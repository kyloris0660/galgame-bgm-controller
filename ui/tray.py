from PIL import Image, ImageDraw
from core.icons import get_exe_icon_pil

try:
    import pystray

    HAVE = True
except Exception:
    HAVE = False


def _default_img():
    img = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((3, 3, 29, 29), outline=(0, 0, 0, 255), width=2)
    d.rectangle((14, 14, 18, 18), fill=(0, 0, 0, 255))
    return img


class Tray:
    def __init__(self, on_select, on_pause, on_quit, on_mode_change, get_mode, on_show):
        self.on_select = on_select
        self.on_pause = on_pause
        self.on_quit = on_quit
        self.on_mode_change = on_mode_change
        self.get_mode = get_mode
        self.on_show = on_show  # 新增：显示主界面
        self.icon = None
        self._img_default = _default_img()
        self._current_img = self._img_default  # 强引用，防止被 GC

    def start(self):
        if not HAVE or self.icon:
            return

        def _mode_label(item):
            return f"静音规则（当前：{self.get_mode()}）"

        def _to_not_fg(icon, item):
            self.on_mode_change("not_foreground")
            try:
                self.icon.update_menu()
            except Exception:
                pass

        def _to_min(icon, item):
            self.on_mode_change("minimized_only")
            try:
                self.icon.update_menu()
            except Exception:
                pass

        menu = pystray.Menu(
            pystray.MenuItem("显示主界面", lambda: self.on_show() or None),  # 新增项
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
                self.icon.visible = False
            except Exception:
                pass
            try:
                self.icon.stop()
            except Exception:
                pass
            self.icon = None
        self._current_img = None  # 允许回收

    def set_icon_from_exe(self, exe_path: str | None):
        if not self.icon:
            return
        try:
            if not exe_path:
                self._current_img = self._img_default
            else:
                self._current_img = get_exe_icon_pil(exe_path, large=True)
            self.icon.icon = self._current_img  # 用强引用
            self.icon.visible = True
        except Exception:
            self._current_img = self._img_default
            try:
                self.icon.icon = self._current_img
                self.icon.visible = True
            except Exception:
                pass
