
from core.policy import should_mute, apply_policy
from core.types import MuteMode, EventSnapshot, TargetStatus

def test_should_mute():
    assert should_mute(MuteMode.NOT_FOREGROUND, is_foreground=False, minimized=False) is True
    assert should_mute(MuteMode.NOT_FOREGROUND, is_foreground=True, minimized=True) is False
    assert should_mute(MuteMode.MINIMIZED_ONLY, is_foreground=True, minimized=True) is True
    assert should_mute(MuteMode.MINIMIZED_ONLY, is_foreground=False, minimized=False) is False

def test_apply_policy():
    snap = EventSnapshot(active_exe="a.exe", targets=[
        TargetStatus(exe="a.exe", is_foreground=True, minimized=False),
        TargetStatus(exe="b.exe", is_foreground=False, minimized=True),
    ])
    act = apply_policy(snap, MuteMode.NOT_FOREGROUND)
    assert act["a.exe"] is False and act["b.exe"] is True
