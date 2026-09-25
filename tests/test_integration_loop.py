
import threading
from core.config import Config
from core.service import Service
from core.audio import DummyAudio
from core.focus import DummyFocus
from core.types import MuteMode

def test_integration_headless(tmp_path):
    cfg = Config(path=str(tmp_path / "cfg.json"))
    cfg.set_targets(["a.exe","b.exe"])
    cfg.set_mode(MuteMode.NOT_FOREGROUND)
    audio = DummyAudio()
    focus = DummyFocus(active_exe="a.exe", minimized=set())
    ready = threading.Event()
    svc = Service(cfg, audio, focus, interval=0.05, on_state=lambda _: ready.set())
    svc.start()
    try:
        assert ready.wait(3), "Service did not publish a snapshot"
        assert audio.state.get("a.exe") is False
        assert audio.state.get("b.exe") is True
    finally:
        svc.stop()
    assert audio.state == {"a.exe": False, "b.exe": False}
