
import os, tempfile
from core.config import Config
from core.types import MuteMode

def test_config_roundtrip(tmp_path):
    cfg = Config(path=os.path.join(tmp_path, "cfg.json"))
    cfg.set_targets(["A.exe","B.exe","A.exe"])
    t = cfg.get_targets()
    assert "A.exe" in t and "B.exe" in t and len(t) == 2
    cfg.set_mode(MuteMode.MINIMIZED_ONLY)
    assert cfg.get_mode() == MuteMode.MINIMIZED_ONLY
