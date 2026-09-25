import argparse
import time
from pathlib import Path

from core.config import Config, DEFAULT_PATH
from core.audio import PycawAudio
from core.focus import WinFocus
from core.instance import SingleInstance
from core.service import Service


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=DEFAULT_PATH)
    args = parser.parse_args()
    instance = SingleInstance(args.config)
    if not instance.primary:
        print("A controller is already running for this configuration.")
        instance.close()
        return
    service = Service(Config(args.config), PycawAudio(journal_path=Path(args.config).with_suffix(".audio-state.json")), WinFocus(),
                      on_error=lambda error: print("Audio:", error or "recovered"))
    try:
        print("BGM Controller running. Press Ctrl+C to restore audio and exit.")
        service.start()
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        service.stop()
        instance.close()


if __name__ == "__main__":
    main()
