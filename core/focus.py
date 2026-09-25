from .paths import exe_key
from .types import EventSnapshot, TargetStatus


class FocusPort:
    def snapshot(self, targets_lower):
        raise NotImplementedError


class DummyFocus(FocusPort):
    def __init__(self, active_exe=None, minimized=None, running=None):
        self.active_exe = active_exe
        self.minimized = minimized or set()
        self.running = running

    def snapshot(self, targets_lower):
        return EventSnapshot(self.active_exe, [
            TargetStatus(t, self.active_exe == t, t in self.minimized)
            for t in targets_lower if self.running is None or t in self.running])


class WinFocus(FocusPort):
    def snapshot(self, targets_lower):
        import psutil
        import win32gui
        import win32process
        from .windows import list_windows

        wanted = set(targets_lower)
        running = set()
        # One process traversal per tick, irrespective of library size.
        for process in psutil.process_iter(["exe"]):
            try:
                key = exe_key(process.info.get("exe") or "")
                if key in wanted:
                    running.add(key)
            except psutil.Error:
                continue
        active = None
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid and not win32gui.IsIconic(hwnd):
                active = exe_key(psutil.Process(pid).exe())
        except (psutil.Error, OSError, win32gui.error):
            pass
        windows = {}
        for window in list_windows():
            windows.setdefault(exe_key(window.exe), []).append(window)
        return EventSnapshot(active, [TargetStatus(
            exe, exe == active,
            bool(windows.get(exe)) and all(w.minimized for w in windows[exe]))
            for exe in targets_lower if exe in running])
