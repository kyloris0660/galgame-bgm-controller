"""Taskbar-style discovery: visible top-level application windows, not processes.

Pinned but closed apps are not tasks. Tool, owned, no-activate, shell and
DWM-cloaked windows are excluded; WS_EX_APPWINDOW can opt an owned window in.
"""
import ctypes
from dataclasses import dataclass
import os

import psutil
import win32con
import win32gui
import win32process

from .paths import exe_key


@dataclass(frozen=True)
class AppWindow:
    hwnd: int
    pid: int
    exe: str
    title: str
    minimized: bool


@dataclass(frozen=True)
class TaskbarApp:
    exe: str
    title: str
    windows: tuple[AppWindow, ...]


def is_taskbar_window(*, visible, owner, style, cloaked, shell=False):
    if not visible or cloaked or shell or style & win32con.WS_EX_TOOLWINDOW:
        return False
    if style & win32con.WS_EX_APPWINDOW:
        return True
    return not owner and not style & win32con.WS_EX_NOACTIVATE


def _cloaked(hwnd):
    value = ctypes.c_int()
    result = ctypes.windll.dwmapi.DwmGetWindowAttribute(
        ctypes.c_void_p(hwnd), 14, ctypes.byref(value), ctypes.sizeof(value))
    return result == 0 and bool(value.value)


def list_windows():
    windows = []
    paths = {}

    def visit(hwnd, _):
        try:
            if not is_taskbar_window(
                visible=win32gui.IsWindowVisible(hwnd),
                owner=win32gui.GetWindow(hwnd, win32con.GW_OWNER),
                style=win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE),
                cloaked=_cloaked(hwnd),
                shell=win32gui.GetClassName(hwnd) in {"Shell_TrayWnd", "Shell_SecondaryTrayWnd", "Progman", "WorkerW"},
            ):
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid == os.getpid():
                return
            if pid not in paths:
                paths[pid] = psutil.Process(pid).exe()
            if paths[pid]:
                windows.append(AppWindow(hwnd, pid, paths[pid], win32gui.GetWindowText(hwnd),
                                         bool(win32gui.IsIconic(hwnd))))
        except (psutil.Error, OSError, win32gui.error):
            return  # A window may disappear while being inspected.

    win32gui.EnumWindows(visit, None)
    return windows


def group_apps(windows):
    grouped = {}
    for window in windows:
        grouped.setdefault(exe_key(window.exe), []).append(window)
    return [TaskbarApp(items[0].exe, items[0].title, tuple(items)) for items in grouped.values()]


def taskbar_apps():
    return group_apps(list_windows())
