import threading
import time
from typing import Optional, Callable, Dict, List
from .config import Config
from .policy import apply_policy
from .audio import AudioPort

try:
    import comtypes

    HAVE_COM = True
except Exception:
    HAVE_COM = False


class Service:
    def __init__(
        self,
        cfg: Config,
        audio: AudioPort,
        focus,
        interval: float = 0.5,
        on_apply: Optional[
            Callable[[Dict[str, bool], List[str], Optional[str]], None]
        ] = None,
    ):
        self.cfg = cfg
        self.audio = audio
        self.focus = focus
        self.interval = interval
        self._stop = threading.Event()
        self._pause = threading.Event()
        self.thread = None
        self.on_apply = on_apply
        self._last_actions: Optional[Dict[str, bool]] = None
        self._last_active: Optional[str] = None

    def start(self):
        if self.thread and self.thread.is_alive():
            return
        self._stop.clear()
        self._pause.clear()
        self.thread = threading.Thread(
            target=self._run, daemon=True, name="BGM-Service"
        )
        self.thread.start()

    def stop(self, wait: float = 3.0):
        self._stop.set()
        self._graceful_unmute_all()
        if self.thread:
            self.thread.join(timeout=wait)

    def pause(self):
        self._pause.set()

    def resume(self):
        self._pause.clear()

    def _run(self):
        com_inited = False
        if HAVE_COM:
            try:
                comtypes.CoInitialize()
                com_inited = True
            except Exception:
                pass
        try:
            while not self._stop.is_set():
                if self._pause.is_set():
                    time.sleep(self.interval)
                    continue
                targets = [t.lower() for t in self.cfg.get_targets()]
                if targets:
                    snap = self.focus.snapshot(targets)
                    mode = self.cfg.get_mode()
                    actions = apply_policy(snap, mode)
                    if (
                        actions != self._last_actions
                        or snap.active_exe != self._last_active
                    ):
                        self.audio.set_bulk(actions)
                        muted = [exe for exe, m in actions.items() if m]
                        if self.on_apply:
                            try:
                                self.on_apply(actions, muted, snap.active_exe)
                            except Exception:
                                pass
                        self._last_actions = dict(actions)
                        self._last_active = snap.active_exe
                time.sleep(self.interval)
        finally:
            if com_inited:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

    def _graceful_unmute_all(self):
        com_inited = False
        if HAVE_COM:
            try:
                comtypes.CoInitialize()
                com_inited = True
            except Exception:
                pass
        try:
            muted_now = []
            if self._last_actions:
                for exe, m in self._last_actions.items():
                    if m:
                        muted_now.append(exe)
            cfg_targets = [t.lower() for t in self.cfg.get_targets()]
            uniq = {}
            for exe in muted_now + cfg_targets:
                uniq[exe] = True
            if hasattr(self.audio, "unmute_multi"):
                self.audio.unmute_multi(list(uniq.keys()))
            else:
                for exe in uniq.keys():
                    try:
                        self.audio.unmute_by_exe(exe)
                    except Exception:
                        pass
        finally:
            if com_inited:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass
