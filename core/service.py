import threading
import time
from .config import Config
from .policy import apply_policy
from .audio import AudioPort
from .focus import FocusPort


class Service:
    """Headless orchestrator; testable without UI."""

    def __init__(
        self, cfg: Config, audio: AudioPort, focus: FocusPort, interval: float = 0.5
    ):
        self.cfg = cfg
        self.audio = audio
        self.focus = focus
        self.interval = interval
        self._stop = threading.Event()
        self._pause = threading.Event()
        self.thread = None

    def start(self):
        if self.thread and self.thread.is_alive():
            return
        self._stop.clear()
        self._pause.clear()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self):
        self._stop.set()

    def pause(self):
        self._pause.set()

    def resume(self):
        self._pause.clear()

    def _run(self):
        while not self._stop.is_set():
            if self._pause.is_set():
                time.sleep(self.interval)
                continue
            targets = [t.lower() for t in self.cfg.get_targets()]
            if targets:
                snapshot = self.focus.snapshot(targets)
                mode = self.cfg.get_mode()
                actions = apply_policy(snapshot, mode)
                self.audio.set_bulk(actions)
            time.sleep(self.interval)
