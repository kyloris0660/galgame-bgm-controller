import hashlib
import pystray

from core.icons import get_exe_icon_pil, placeholder
from .theme import MODE_LABELS


def state_label(state):
    if not state.controlled:
        return "本次已停止"
    return "已暂停" if state.paused else ("自动静音" if state.muted else "允许播放")


class Tray:
    """One icon per running registered executable, plus an idle fallback."""

    def __init__(self, on_select, on_pause, on_quit, on_mode_change, on_show,
                 on_target_pause, get_name, icon_factory=pystray.Icon, on_target_stop=None):
        self.on_select, self.on_pause, self.on_quit = on_select, on_pause, on_quit
        self.on_mode_change, self.on_show = on_mode_change, on_show
        self.on_target_pause, self.get_name = on_target_pause, get_name
        self.on_target_stop = on_target_stop
        self.icon_factory = icon_factory
        self.icon = None
        self.games = {}
        self.states = {}
        self.paused = False

    def _common(self):
        item = pystray.MenuItem
        return [
            item("打开游戏库", self.on_show, default=True),
            item("添加游戏…", self.on_select),
            pystray.Menu.SEPARATOR,
            item("继续全部" if self.paused else "暂停全部", self.on_pause),
            item("退出全部控制器", self.on_quit),
        ]

    def start(self):
        if not self.icon:
            self.icon = self.icon_factory("bgm_library", placeholder(), "Galgame BGM · 等待游戏启动",
                                          pystray.Menu(*self._common()))
            self.icon.run_detached()

    def _menu(self, exe, state):
        item = pystray.MenuItem

        def toggle():
            self.on_target_pause(exe)

        def background():
            self.on_mode_change("not_foreground", exe)

        def minimized():
            self.on_mode_change("minimized_only", exe)

        def inherit():
            self.on_mode_change(None, exe)

        def stop_target():
            if self.on_target_stop:
                self.on_target_stop(exe)

        return pystray.Menu(
            item(self.get_name(exe)[:60], None, enabled=False),
            item(state_label(state), None, enabled=False),
            item("继续此游戏" if state.paused else "暂停此游戏", toggle, enabled=not self.paused),
            item("停止此游戏控制（本次）", stop_target, enabled=self.on_target_stop is not None),
            item("静音规则", pystray.Menu(
                item(MODE_LABELS["not_foreground"], background, checked=lambda _: state.mode.value == "not_foreground"),
                item(MODE_LABELS["minimized_only"], minimized, checked=lambda _: state.mode.value == "minimized_only"),
                item("跟随默认规则", inherit))),
            pystray.Menu.SEPARATOR, *self._common())

    def sync(self, states, paused=False):
        states = {exe: state for exe, state in states.items() if state.controlled}
        self.start()
        pause_changed = self.paused != paused
        self.paused = paused
        # Create replacements before removing the last reachable icon.
        for exe, state in states.items():
            if exe not in self.games:
                name = "bgm_" + hashlib.sha256(exe.encode()).hexdigest()[:16]
                icon = self.icon_factory(name, get_exe_icon_pil(exe),
                                         f"{self.get_name(exe)[:70]} · {state_label(state)}",
                                         self._menu(exe, state))
                icon.run_detached()
                self.games[exe] = icon
            elif self.states.get(exe) != state or pause_changed:
                icon = self.games[exe]
                icon.title = f"{self.get_name(exe)[:70]} · {state_label(state)}"
                icon.menu = self._menu(exe, state)
        self.icon.visible = not bool(states)
        self.icon.menu = pystray.Menu(*self._common())
        for exe in set(self.games) - set(states):
            self.games.pop(exe).stop()
        self.states = dict(states)

    def stop(self):
        for icon in list(self.games.values()) + ([self.icon] if self.icon else []):
            icon.stop()
        self.games.clear()
        self.icon = None
        self.states.clear()
