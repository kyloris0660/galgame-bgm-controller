from types import SimpleNamespace
import psutil
import win32gui
import win32process
from core.focus import WinFocus
from core.windows import AppWindow
from core.paths import exe_key


def test_single_process_scan_and_multiple_windows_minimized(monkeypatch):
    import core.windows
    a, b = exe_key(r"C:\a.exe"), exe_key(r"C:\b.exe")
    calls = []
    def processes(attrs):
        calls.append(attrs)
        return [SimpleNamespace(info={"exe": a}), SimpleNamespace(info={"exe": b})]
    monkeypatch.setattr(psutil, "process_iter", processes)
    monkeypatch.setattr(win32gui, "GetForegroundWindow", lambda: 20)
    monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda _: (1, 2))
    monkeypatch.setattr(win32gui, "IsIconic", lambda _: False)
    monkeypatch.setattr(psutil, "Process", lambda _: SimpleNamespace(exe=lambda: b))
    monkeypatch.setattr(core.windows, "list_windows", lambda: [
        AppWindow(10, 1, a, "A", True), AppWindow(11, 1, a, "A dialog", False),
        AppWindow(20, 2, b, "B", False)])
    snap = WinFocus().snapshot([a, b, "missing.exe"])
    assert len(calls) == 1
    assert len(snap.targets) == 2
    assert not snap.targets[0].minimized
    assert snap.targets[1].is_foreground
