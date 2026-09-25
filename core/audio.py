"""All-render-device audio control with exact-path identity and a recovery journal."""
from dataclasses import dataclass, field
import json
import logging
import os
from pathlib import Path
import tempfile

from .paths import exe_key

log = logging.getLogger(__name__)


@dataclass
class SessionBatch:
    sessions: list
    endpoints: set = field(default_factory=set)
    errors: list = field(default_factory=list)


class DeviceSession:
    def __init__(self, endpoint, session):
        self.endpoint = endpoint
        self.session = session

    def __getattr__(self, name):
        return getattr(self.session, name)


def scan_sessions():
    """Rescan active output devices, including non-default outputs, every tick.

    Polling also covers default-device switches, unplug/replug and audio-service
    restarts without retaining an invalid device enumerator.
    """
    import comtypes
    from pycaw.pycaw import AudioUtilities
    from pycaw.api.audiopolicy import IAudioSessionManager2, IAudioSessionControl2
    from pycaw.constants import EDataFlow, DEVICE_STATE
    from pycaw.utils import AudioSession

    enumerator = AudioUtilities.GetDeviceEnumerator()
    collection = enumerator.EnumAudioEndpoints(EDataFlow.eRender.value, DEVICE_STATE.ACTIVE.value)
    result = SessionBatch([])
    for index in range(collection.GetCount()):
        try:
            device = collection.Item(index)
            endpoint = device.GetId()
            interface = device.Activate(IAudioSessionManager2._iid_, comtypes.CLSCTX_ALL, None)
            manager = interface.QueryInterface(IAudioSessionManager2)
            sessions = manager.GetSessionEnumerator()
            # Mark complete only after this endpoint was successfully enumerated.
            current = []
            for session_index in range(sessions.GetCount()):
                control = sessions.GetSession(session_index).QueryInterface(IAudioSessionControl2)
                current.append(DeviceSession(endpoint, AudioSession(control)))
            result.sessions.extend(current)
            result.endpoints.add(endpoint)
        except Exception as exc:
            result.errors.append(str(exc))
    return result


class AudioPort:
    def set_bulk(self, mutes):
        raise NotImplementedError

    def restore_all(self):
        self.set_bulk({})

    def release(self):
        pass


class DummyAudio(AudioPort):
    def __init__(self):
        self.state = {}
        self.history = []

    def set_bulk(self, mutes):
        self.state.update({key: False for key in self.state if key not in mutes})
        self.state.update(mutes)
        self.history.append(dict(mutes))


class PycawAudio(AudioPort):
    def __init__(self, sessions=None, journal_path=None):
        self._sessions = sessions or scan_sessions
        self._owned = {}
        self._pending = {}
        self.journal_path = Path(journal_path) if journal_path else None
        if self.journal_path and self.journal_path.exists():
            self._pending = json.loads(self.journal_path.read_text(encoding="utf-8"))

    @staticmethod
    def _key(session):
        return json.dumps([getattr(session, "endpoint", "default"), session.InstanceIdentifier])

    def _save(self):
        if not self.journal_path:
            return
        self.journal_path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(dir=self.journal_path.parent, prefix=".bgm-audio-")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as stream:
                json.dump(self._pending, stream, ensure_ascii=False)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.journal_path)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)

    def _forget(self, key):
        self._owned.pop(key, None)
        self._pending.pop(key, None)
        self._save()

    def _restore(self, key):
        session, volume = self._owned[key]
        try:
            volume.SetMute(self._pending[key]["mute"], None)
        except Exception:
            # Expired sessions no longer have audio to restore.
            if session.State != 2:
                raise
        self._forget(key)

    def set_bulk(self, mutes):
        wanted = {exe_key(exe) for exe, mute in mutes.items() if mute}
        batch = self._sessions()
        if not isinstance(batch, SessionBatch):
            batch = SessionBatch(list(batch), {"default"})
        seen = set()
        failures = list(batch.errors)
        for session in batch.sessions:
            try:
                key = self._key(session)
                seen.add(key)
                process = session.Process
                path = exe_key(process.exe()) if process else ""
                if key in self._pending:
                    # Refresh stale COM handles after an endpoint reconnects.
                    self._owned[key] = (session, session.SimpleAudioVolume)
                    if path != self._pending[key]["exe"]:
                        continue
                if path not in wanted:
                    if key in self._owned:
                        self._restore(key)
                    continue
                if session.State == 2:
                    if key in self._owned:
                        self._restore(key)
                    continue
                volume = session.SimpleAudioVolume
                if key not in self._pending:
                    self._pending[key] = {"exe": path, "mute": bool(volume.GetMute())}
                    try:
                        self._save()  # Persist baseline before mutating Core Audio.
                    except Exception:
                        self._pending.pop(key, None)
                        raise
                self._owned[key] = (session, volume)
                if not volume.GetMute():
                    volume.SetMute(True, None)
            except Exception as exc:
                failures.append(str(exc))
        for key in list(self._pending):
            if key in seen:
                continue
            endpoint, _ = json.loads(key)
            if endpoint in batch.endpoints:
                # Restore a still-accessible old handle before pruning an ended
                # instance; enumeration can race with session shutdown.
                if key in self._owned:
                    try:
                        self._restore(key)
                    except Exception as exc:
                        failures.append(str(exc))
                else:
                    self._forget(key)
            elif key in self._owned:
                try:
                    self._restore(key)
                except Exception:
                    pass  # Offline endpoint: keep the baseline for reconnection.
        if failures:
            raise RuntimeError("部分音频设备或会话暂时不可用，将自动重试：" + failures[0])

    def restore_all(self):
        # Rebind pending records first, including records from a previous crash.
        try:
            self.set_bulk({})
        except Exception:
            log.exception("Some sessions could not be rebound for restoration")
        for key in list(self._owned):
            try:
                self._restore(key)
            except Exception:
                log.warning("Restoration deferred until the audio device returns", exc_info=True)
        if self._pending:
            log.warning("%s audio baselines remain in the recovery journal", len(self._pending))

    def release(self):
        # COM references must be released on their owning worker, before CoUninitialize.
        self._owned.clear()
