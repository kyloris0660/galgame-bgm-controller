from typing import List, Optional, Set
from .types import EventSnapshot, TargetStatus


class FocusPort:
    """Interface for foreground/minimized detection."""

    def snapshot(self, targets_lower: List[str]) -> EventSnapshot: ...


class DummyFocus(FocusPort):
    def __init__(self, active_exe: Optional[str], minimized: Set[str] = set()):
        self.active_exe = active_exe
        self.minimized = minimized

    def snapshot(self, targets_lower: List[str]) -> EventSnapshot:
        tstats = []
        for t in targets_lower:
            tstats.append(
                TargetStatus(
                    exe=t,
                    is_foreground=(self.active_exe == t),
                    minimized=(t in self.minimized),
                )
            )
        return EventSnapshot(active_exe=self.active_exe, targets=tstats)


try:
    import psutil
    import win32gui
    import win32process

    HAVE_WIN = True
except Exception:
    HAVE_WIN = False


class WinFocus(FocusPort):
    def _get_foreground(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            exe = None
            if pid:
                try:
                    p = psutil.Process(pid)
                    exe = (p.exe() or "").lower()
                except Exception:
                    exe = None
            minimized = bool(win32gui.IsIconic(hwnd))
            return exe, minimized
        except Exception:
            return None, False

    def _pids_for_exe(self, exe_lower: str):
        ids = set()
        for p in psutil.process_iter(["exe", "pid"]):
            try:
                if (p.info.get("exe") or "").lower() == exe_lower:
                    ids.add(int(p.info["pid"]))
            except Exception:
                pass
        return ids

    def _is_process_minimized(self, pids: Set[int]) -> bool:
        if not pids or not HAVE_WIN:
            return False
        found_window = False
        minimized_all = True

        def _enum_cb(hwnd, _):
            nonlocal found_window, minimized_all
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid in pids:
                    found_window = True
                    if not win32gui.IsIconic(hwnd):
                        minimized_all = False
            except Exception:
                pass

        try:
            win32gui.EnumWindows(_enum_cb, None)
        except Exception:
            return False
        if not found_window:
            return False
        return minimized_all

    def snapshot(self, targets_lower: List[str]) -> EventSnapshot:
        active_exe, fg_min = self._get_foreground()
        tstats = []
        for t in targets_lower:
            pids = self._pids_for_exe(t)
            minimized = self._is_process_minimized(pids)
            is_fg = (active_exe == t) and not fg_min
            tstats.append(TargetStatus(exe=t, is_foreground=is_fg, minimized=minimized))
        return EventSnapshot(active_exe=active_exe, targets=tstats)
