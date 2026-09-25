"""Verify a built GUI using isolated silent fixtures and an isolated config."""
import json
from pathlib import Path
import subprocess
import sys
import time

from windows_audio_smoke import ROOT, wait_for
from core.audio import scan_sessions
from core.config import Config
from core.paths import exe_key


def main():
    exe = str(Path(sys.argv[1]).resolve())
    folder = ROOT / ".artifacts"
    token = str(time.time_ns())
    games = [str(folder / f"Game{x}.exe") for x in "AB"]
    stops = [folder / f"package-stop-{x}-{token}" for x in "AB"]
    children = []
    controller = None
    cfg = Config(str(folder / f"package-{token}.json"))
    cfg.set_targets(games)
    def sessions():
        result = {}
        for s in scan_sessions().sessions:
            try:
                if s.Process and exe_key(s.Process.exe()) in {exe_key(g) for g in games} and s.State != 2:
                    result[exe_key(s.Process.exe())] = s.SimpleAudioVolume
            except Exception:
                pass
        return result
    try:
        for index, game in enumerate(games):
            children.append(subprocess.Popen([game, f"Packaged BGM Test {index}", str(stops[index])]))
        wait_for(lambda: len(sessions()) == 2)
        for volume in sessions().values():
            volume.SetMute(False, None)
        controller = subprocess.Popen([exe, "--config", cfg.path, "--smoke-test", "8"])
        wait_for(lambda: len(sessions()) == 2 and all(v.GetMute() for v in sessions().values()))
        duplicate = subprocess.Popen([exe, "--config", cfg.path])
        assert duplicate.wait(timeout=6) == 0
        assert controller.poll() is None
        assert controller.wait(timeout=15) == 0
        assert not any(v.GetMute() for v in sessions().values())
        journal = Path(cfg.path).with_suffix(".audio-state.json")
        assert json.loads(journal.read_text()) == {}
        log = Path(cfg.path).with_suffix(".log")
        assert not log.read_text(encoding="utf-8").strip()
        print(json.dumps({"executable": exe, "two_real_sessions_muted": True,
                          "duplicate_exits_without_second_controller": True,
                          "graceful_exit_restores_both": True, "journal_empty": True,
                          "error_log_empty": True}, indent=2))
    finally:
        if controller and controller.poll() is None:
            # A smoke-test instance normally exits itself; wait before any cleanup.
            controller.wait(timeout=20)
        for volume in sessions().values():
            volume.SetMute(False, None)  # Only isolated fixture paths.
        for stop in stops:
            stop.touch()
        for child in children:
            child.wait(timeout=5)


if __name__ == "__main__":
    main()
