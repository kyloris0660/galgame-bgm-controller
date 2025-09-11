# core/service.py
import threading
import time
from typing import Optional, Callable, Dict
from .config import Config
from .policy import apply_policy
from .audio import AudioPort
from . import debug as dbg

try:
    import comtypes

    HAVE_COM = True
except Exception:
    HAVE_COM = False


class Service:
    """读取状态 -> 策略 -> 执行静音；变化才回调 on_apply(actions, muted_list, active_exe)"""

    def __init__(
        self,
        cfg: Config,
        audio: AudioPort,
        focus,
        interval: float = 0.5,
        on_apply: Optional[
            Callable[[Dict[str, bool], list, Optional[str]], None]
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
        self._last_actions = None
        self._last_active = None

    def start(self):
        if self.thread and self.thread.is_alive():
            return
        self._stop.clear()
        self._pause.clear()
        self.thread = threading.Thread(
            target=self._run, daemon=True, name="BGM-Service"
        )
        self.thread.start()

    def stop(self):
        self._stop.set()

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
                    # 变化才执行/回调
                    if (
                        actions != self._last_actions
                        or snap.active_exe != self._last_active
                    ):
                        self.audio.set_bulk(actions)
                        muted = [exe for exe, m in actions.items() if m]
                        dbg.log_change(
                            "apply",
                            {
                                "mode": mode.value,
                                "active": snap.active_exe,
                                "actions": actions,
                            },
                        )
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
