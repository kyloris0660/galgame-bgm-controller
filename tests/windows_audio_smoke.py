"""Opt-in real Windows test against two silent fixture processes only.

Compile tests/fixtures/AudioFixture.cs to .artifacts/GameA.exe and GameB.exe.
Run: python tests/windows_audio_smoke.py
No user configuration, default device, master volume or other sessions are changed.
"""
import json
from pathlib import Path
import subprocess
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from core.audio import PycawAudio, scan_sessions
from core.config import Config
from core.focus import WinFocus
from core.paths import exe_key
from core.service import Service
from core.types import MuteMode
from ui.tray import Tray


def wait_for(predicate, timeout=8):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        result = predicate()
        if result:
            return result
        time.sleep(.1)
    raise AssertionError("Timed out waiting for fixture")


def main():
    folder = ROOT / ".artifacts"
    token = str(time.time_ns())
    paths = [str(folder / f"Game{letter}.exe") for letter in "AB"]
    stops = [folder / f"stop-{letter}-{token}" for letter in "AB"]
    children = []
    cfg = Config(str(folder / f"native-{token}.json"))
    cfg.set_targets(paths)
    cfg.set_mode(MuteMode.NOT_FOREGROUND)
    audio = PycawAudio(journal_path=folder / f"native-{token}.audio-state.json")
    states = []
    tray = Tray(lambda: None, lambda: None, lambda: None, lambda *_: None,
                lambda: None, lambda _: None, cfg.get_name)
    service = Service(cfg, audio, WinFocus(), on_state=lambda s: (states.append(s), tray.sync(s)))
    checks = []
    def volumes():
        found = {}
        for s in scan_sessions().sessions:
            try:
                if s.Process and exe_key(s.Process.exe()) in {exe_key(p) for p in paths} and s.State != 2:
                    found[exe_key(s.Process.exe())] = bool(s.SimpleAudioVolume.GetMute())
            except Exception:
                pass
        return found
    def initialize_fixture(path):
        # Windows persists mixer settings between fixture runs. Set a known
        # baseline only for this isolated test executable, never for user apps.
        for session in scan_sessions().sessions:
            if session.Process and exe_key(session.Process.exe()) == exe_key(path) and session.State != 2:
                session.SimpleAudioVolume.SetMute(False, None)
    try:
        children.append(subprocess.Popen([paths[0], "BGM Test A", str(stops[0])]))
        wait_for(lambda: exe_key(paths[0]) in volumes())
        initialize_fixture(paths[0])
        service.tick()
        assert len(tray.games) == 1
        checks.append("one running game creates one native tray icon")
        children.append(subprocess.Popen([paths[1], "BGM Test B", str(stops[1])]))
        wait_for(lambda: len(volumes()) == 2)
        initialize_fixture(paths[1])
        service.tick()
        assert len(tray.games) == 2
        assert states[-1][exe_key(paths[0])].muted  # B is foreground after startup.
        assert volumes()[exe_key(paths[0])]
        checks.append("late second game creates a second native icon; background audio muted")
        service.toggle_target(paths[0])
        service.tick()
        assert not volumes()[exe_key(paths[0])]
        checks.append("pausing A restores only A")
        service.toggle_target(paths[0])
        service.tick()
        service.dismiss_target(paths[0])
        service.tick()
        assert len(tray.games) == 1 and exe_key(paths[1]) in tray.games
        assert not volumes()[exe_key(paths[0])]
        checks.append("stopping A controller leaves B icon and audio control alive")
        service.toggle_target(paths[0])
        service.tick()
        stops[0].touch()
        children[0].wait(timeout=5)
        service.tick()
        assert set(tray.games) == {exe_key(paths[1])}
        checks.append("closing A leaves B controlled")
        restart_stop = folder / f"stop-A-restarted-{token}"
        stops.append(restart_stop)
        children.append(subprocess.Popen([paths[0], "BGM Test A restarted", str(restart_stop)]))
        wait_for(lambda: len(volumes()) == 2)
        assert not volumes()[exe_key(paths[0])], "Closed-game mute leaked into the next launch"
        service.tick()
        assert len(tray.games) == 2
        checks.append("restarting A automatically restores its icon without a leaked mute")
        service.pause()
        service.tick()
        assert not volumes()[exe_key(paths[1])]
        checks.append("global pause restores remaining session")
        service.resume()
        service.tick()
        audio.restore_all()
        assert not volumes()[exe_key(paths[1])]
        checks.append("graceful shutdown restores remaining session")
        print(json.dumps({"passed": checks, "native_icons": True, "silent_fixture_audio": True}, indent=2))
    finally:
        audio.restore_all()
        tray.stop()
        for stop in stops:
            stop.touch()
        for child in children:
            child.wait(timeout=5)
        audio.release()


if __name__ == "__main__":
    main()
