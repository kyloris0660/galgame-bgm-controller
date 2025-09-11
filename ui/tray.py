from PIL import Image, ImageDraw

try:
    import pystray

    HAVE_PYSTRAY = True
except Exception:
    HAVE_PYSTRAY = False


def _default_image():
    img = Image.new("RGBA", (64, 64), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((8, 8, 56, 56), outline=(0, 0, 0), width=3)
    d.rectangle((22, 22, 42, 42), fill=(0, 0, 0))
    return img


class Tray:
    def __init__(self, on_select, on_pause, on_quit, on_mode_change, get_mode):
        self.on_select = on_select
        self.on_pause = on_pause
        self.on_quit = on_quit
        self.on_mode_change = on_mode_change
        self.get_mode = get_mode
        self.icon = None
        self.img = _default_image()

    def start(self):
        if not HAVE_PYSTRAY:
            return
        if self.icon:
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
        self.icon = pystray.Icon("gal_bgm_controller", self.img, "BGM Controller", menu)
        self.icon.run_detached()

    def stop(self):
        if self.icon:
            try:
                self.icon.stop()
            except Exception:
                pass

    def update_icon_from_exe(self, exe_path: str):
        # 简化：可拓展为真实图标
        pass

    def reset_icon_default(self):
        pass

    def notify(self, text: str):
        pass
