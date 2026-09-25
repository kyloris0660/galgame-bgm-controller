import threading
from core.audio import DummyAudio
from core.config import Config
from core.focus import DummyFocus
from core.service import Service
from core.types import MuteMode


def setup_service(tmp_path, running=None):
    config = Config(str(tmp_path / "config.json"))
    config.set_targets(["a.exe", "b.exe"])
    audio = DummyAudio()
    focus = DummyFocus(active_exe="a.exe", running=running)
    observed = []
    service = Service(config, audio, focus, on_state=observed.append)
    return config, audio, focus, service, observed


def test_close_restart_pause_and_remove_are_independent(tmp_path):
    cfg, audio, focus, service, observed = setup_service(tmp_path, {"a.exe", "b.exe"})
    service.tick()
    assert audio.state == {"a.exe": False, "b.exe": True}
    focus.running = {"b.exe"}
    service.tick()
    assert set(observed[-1]) == {"b.exe"}
    assert audio.state["b.exe"] is True
    focus.running.add("a.exe")
    focus.active_exe = None
    service.tick()
    assert set(observed[-1]) == {"a.exe", "b.exe"}
    service.toggle_target("a.exe")
    service.tick()
    assert audio.state == {"a.exe": False, "b.exe": True}
    assert observed[-1]["a.exe"].paused
    cfg.set_targets(["a.exe"])
    service.tick()
    assert audio.state["b.exe"] is False
    service.pause()
    service.tick()
    assert all(not value for value in audio.state.values())
    service.resume()
    service.toggle_target("a.exe")
    service.tick()
    assert audio.state["a.exe"] is True


def test_per_game_mode_and_no_policy_change_still_rescans(tmp_path):
    cfg, audio, focus, service, _ = setup_service(tmp_path)
    cfg.set_mode(MuteMode.MINIMIZED_ONLY, "b.exe")
    service.tick()
    assert audio.state["b.exe"] is False
    focus.minimized.add("b.exe")
    service.tick()
    assert audio.state["b.exe"] is True
    count = len(audio.history)
    service.tick()
    assert len(audio.history) == count + 1
    cfg.set_mode(None, "b.exe")
    assert cfg.get_mode("b.exe") == MuteMode.NOT_FOREGROUND


def test_audio_failure_does_not_freeze_game_lifecycle(tmp_path):
    _, audio, focus, service, observed = setup_service(tmp_path, {"a.exe", "b.exe"})
    service.tick()
    focus.running = {"b.exe"}
    def fail(_):
        raise OSError("device invalidated")
    audio.set_bulk = fail
    import pytest
    with pytest.raises(OSError):
        service.tick()
    assert set(observed[-1]) == {"b.exe"}


def test_stop_one_controller_resumes_on_next_game_launch(tmp_path):
    _, audio, focus, service, observed = setup_service(tmp_path, {"a.exe", "b.exe"})
    focus.active_exe = None
    service.tick()
    service.dismiss_target("a.exe")
    service.tick()
    assert not observed[-1]["a.exe"].controlled
    assert audio.state == {"a.exe": False, "b.exe": True}
    focus.running.remove("a.exe")
    service.tick()
    focus.running.add("a.exe")
    service.tick()
    assert observed[-1]["a.exe"].controlled
    assert audio.state == {"a.exe": True, "b.exe": True}


def test_worker_retries_after_device_failure(tmp_path):
    _, audio, _, service, _ = setup_service(tmp_path)
    recovered = threading.Event()
    original = audio.set_bulk
    calls = []
    def unreliable(actions):
        calls.append(actions)
        if len(calls) == 1:
            raise OSError("device invalidated")
        original(actions)
    audio.set_bulk = unreliable
    service.interval = .01
    errors = []
    def on_error(error):
        errors.append(error)
        if error is None:
            recovered.set()
    service.on_error = on_error
    service.start()
    try:
        assert recovered.wait(3)
        assert errors == ["device invalidated", None]
    finally:
        service.stop()
    assert audio.state["b.exe"] is False
