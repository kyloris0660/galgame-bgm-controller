from core.service import GameState
from core.types import MuteMode
from ui.tray import Tray


class Icon:
    def __init__(self, name, image, title, menu):
        self.name, self.title, self.menu = name, title, menu
        self.visible = False
        self.started = self.stopped = 0

    def run_detached(self):
        self.started += 1
        self.visible = True

    def stop(self):
        self.stopped += 1
        self.visible = False


def test_each_game_has_independent_icon_and_idle_fallback():
    tray = Tray(lambda: None, lambda: None, lambda: None, lambda *_: None,
                lambda: None, lambda _: None, lambda x: x, icon_factory=Icon)
    a = GameState("a.exe", False, False, MuteMode.NOT_FOREGROUND)
    b = GameState("b.exe", True, False, MuteMode.MINIMIZED_ONLY)
    tray.sync({})
    assert tray.icon.visible
    tray.sync({"a.exe": a, "b.exe": b})
    icon_a, icon_b = tray.games["a.exe"], tray.games["b.exe"]
    assert icon_a is not icon_b and not tray.icon.visible
    tray.sync({"b.exe": b})
    assert icon_a.stopped == 1 and icon_b.stopped == 0
    assert tray.games["b.exe"] is icon_b
    tray.sync({"a.exe": a, "b.exe": b})
    assert tray.games["b.exe"] is icon_b
    assert tray.games["a.exe"] is not icon_a
    tray.sync({})
    assert tray.icon.visible and icon_b.stopped == 1
    tray.stop()
    assert tray.icon is None and not tray.games
