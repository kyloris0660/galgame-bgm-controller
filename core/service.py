from dataclasses import dataclass
import logging
import threading

from .paths import exe_key
from .policy import should_mute
from .types import MuteMode

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class GameState:
    exe: str
    muted: bool
    paused: bool
    mode: MuteMode
    controlled: bool = True


class Service:
    """One worker owns Core Audio COM objects; UI receives value snapshots."""

    def __init__(self, cfg, audio, focus, interval=0.5, on_state=None, on_error=None):
        self.cfg, self.audio, self.focus = cfg, audio, focus
        self.interval = interval
        self.on_state, self.on_error = on_state, on_error
        self._stop = threading.Event()
        self._wake = threading.Event()
        self._lock = threading.Lock()
        self._paused = False
        self._paused_targets = set()
        self._dismissed = set()
        self.thread = None
        self._last_states = None
        self.error = None

    @property
    def paused(self):
        with self._lock:
            return self._paused

    def toggle_target(self, exe):
        with self._lock:
            key = exe_key(exe)
            if key in self._dismissed:
                self._dismissed.remove(key)
                self._paused_targets.discard(key)
            else:
                self._paused_targets.symmetric_difference_update({key})
        self.wake()

    def dismiss_target(self, exe):
        with self._lock:
            self._dismissed.add(exe_key(exe))
        self.wake()

    def pause(self):
        with self._lock:
            self._paused = True
        self.wake()

    def resume(self):
        with self._lock:
            self._paused = False
        self.wake()

    def wake(self):
        self._wake.set()

    def start(self):
        if self.thread and self.thread.is_alive():
            return
        self._stop.clear()
        self.thread = threading.Thread(target=self._run, name="BGM-Service", daemon=True)
        self.thread.start()

    def stop(self, wait=5.0):
        self._stop.set()
        self.wake()
        if self.thread:
            self.thread.join(timeout=wait)
            if self.thread.is_alive():
                raise RuntimeError("音频控制线程仍在退出，请稍后重试")
        if self.error:
            raise RuntimeError(self.error)

    def tick(self):
        targets = [exe_key(t) for t in self.cfg.get_targets()]
        snapshot = self.focus.snapshot(targets)
        with self._lock:
            self._paused_targets.intersection_update(targets)
            self._dismissed.intersection_update(t.exe for t in snapshot.targets)
            paused_all, paused_targets = self._paused, self._paused_targets.copy()
            dismissed = self._dismissed.copy()
        states, actions = {}, {}
        for target in snapshot.targets:
            controlled = target.exe not in dismissed
            paused = paused_all or target.exe in paused_targets or not controlled
            mode = self.cfg.get_mode(target.exe)
            mute = not paused and should_mute(mode, target.is_foreground, target.minimized)
            states[target.exe] = GameState(target.exe, mute, paused, mode, controlled)
            if not paused:
                actions[target.exe] = mute
        try:
            self.audio.set_bulk(actions)
        finally:
            # A failing/disconnected device must not freeze other games' lifecycle.
            if states != self._last_states:
                if self.on_state:
                    self.on_state(states)
                self._last_states = states

    def _report(self, error):
        if self.error != error:
            self.error = error
            if self.on_error:
                self.on_error(error)

    def _run(self):
        import comtypes
        comtypes.CoInitialize()
        try:
            while not self._stop.is_set():
                self._wake.clear()
                try:
                    self.tick()
                    self._report(None)
                except Exception as exc:
                    if str(exc) != self.error:
                        log.exception("Controller tick failed; retrying")
                    self._report(str(exc))
                self._wake.wait(self.interval)
        finally:
            try:
                self.audio.restore_all()
                self._report(None)
            except Exception as exc:
                self._report(str(exc))
            finally:
                self.audio.release()
                comtypes.CoUninitialize()
