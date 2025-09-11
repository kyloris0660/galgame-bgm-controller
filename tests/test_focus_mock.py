
from core.focus import DummyFocus
from core.types import EventSnapshot

def test_focus_dummy():
    f = DummyFocus(active_exe="a.exe", minimized={"b.exe"})
    snap = f.snapshot(["a.exe","b.exe"])
    t = {ts.exe: (ts.is_foreground, ts.minimized) for ts in snap.targets}
    assert t["a.exe"] == (True, False)
    assert t["b.exe"] == (False, True)
