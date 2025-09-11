
import time
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
    svc = Service(cfg, audio, focus, interval=0.05)
    svc.start()
    time.sleep(0.15)
    svc.stop()
    # a 前台 -> 不静音；b 非前台 -> 静音
    assert audio.state.get("a.exe") is False
    assert audio.state.get("b.exe") is True
