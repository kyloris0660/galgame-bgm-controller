import win32con
import pytest
from core.windows import is_taskbar_window, group_apps, AppWindow


@pytest.mark.parametrize("changes,expected", [
    ({}, True), ({"visible": False}, False), ({"cloaked": True}, False),
    ({"owner": 7}, False), ({"style": win32con.WS_EX_TOOLWINDOW}, False),
    ({"style": win32con.WS_EX_NOACTIVATE}, False),
    ({"owner": 7, "style": win32con.WS_EX_APPWINDOW}, True),
    ({"shell": True}, False),
])
def test_taskbar_window_rules(changes, expected):
    properties = dict(visible=True, owner=0, style=0, cloaked=False)
    properties.update(changes)
    assert bool(is_taskbar_window(**properties)) is expected


def test_multiwindow_same_executable_is_one_card_and_minimized_is_kept():
    windows = [AppWindow(1, 10, r"C:\Game.exe", "Game", False),
               AppWindow(2, 10, r"c:\GAME.exe", "Settings", True),
               AppWindow(3, 11, r"D:\Game.exe", "Other game", True)]
    apps = group_apps(windows)
    assert len(apps) == 2
    assert len(apps[0].windows) == 2
    assert apps[1].windows[0].minimized
