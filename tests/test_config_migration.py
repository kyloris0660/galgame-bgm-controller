import json
import pytest
from core.config import Config
from core.types import MuteMode


def test_legacy_preserved_until_user_changes_and_case_deduplicated(tmp_path):
    path = tmp_path / "old.json"
    raw = json.dumps({"selected_procs": [r"C:\Game.exe", r"c:\GAME.exe"],
                      "mute_mode": "minimized_only", "unknown": {"keep": 1}})
    path.write_text(raw, encoding="utf-8")
    cfg = Config(str(path))
    assert path.read_text(encoding="utf-8") == raw
    assert len(cfg.get_targets()) == 1
    assert cfg.get_mode() == MuteMode.MINIMIZED_ONLY
    cfg.set_name(r"C:\Game.exe", "游戏")
    assert cfg.read()["unknown"] == {"keep": 1}
    assert Config(str(path)).get_name(r"c:\GAME.exe") == "游戏"


def test_corrupt_config_not_overwritten(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text("{broken", encoding="utf-8")
    with pytest.raises(ValueError):
        Config(str(path))
    assert path.read_text() == "{broken"


def test_atomic_write_failure_preserves_disk_and_memory(tmp_path, monkeypatch):
    path = tmp_path / "settings.json"
    cfg = Config(str(path))
    cfg.set_targets(["a.exe"])
    original = path.read_bytes()
    def fail(*_):
        raise PermissionError("locked")
    monkeypatch.setattr("core.config.os.replace", fail)
    with pytest.raises(PermissionError):
        cfg.set_targets(["b.exe"])
    assert path.read_bytes() == original
    assert cfg.get_targets() == ["a.exe"]
