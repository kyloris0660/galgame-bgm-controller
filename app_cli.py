
import time, sys
from core.config import Config
from core.audio import PycawAudio, DummyAudio
from core.focus import WinFocus, DummyFocus
from core.service import Service
from core.types import MuteMode

def main():
    cfg = Config()
    # Choose real or dummy based on availability
    audio = PycawAudio() if 'pycaw' in sys.modules else PycawAudio()
    focus = WinFocus()
    svc = Service(cfg, audio, focus, interval=0.5)
    print("[CLI] Running headless controller. Press Ctrl+C to stop.")
    print(f"[CLI] Targets: {cfg.get_targets()}  Mode: {cfg.get_mode().value}")
    try:
        svc.start()
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[CLI] Stopping...")
    finally:
        svc.stop()

if __name__ == "__main__":
    main()
