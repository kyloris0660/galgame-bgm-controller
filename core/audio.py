class AudioPort:
    """Interface for audio control; real impl uses pycaw, test uses Dummy."""

    def mute(self, exe_path: str): ...
    def unmute(self, exe_path: str): ...
    def set_bulk(self, mutes: dict):
        """mutes: exe->bool"""
        for exe, flag in mutes.items():
            if flag:
                self.mute(exe)
            else:
                self.unmute(exe)


class DummyAudio(AudioPort):
    def __init__(self):
        self.state = {}  # exe->muted?

    def mute(self, exe_path: str):
        self.state[exe_path] = True

    def unmute(self, exe_path: str):
        self.state[exe_path] = False


try:
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
    from comtypes import CLSCTX_ALL  # noqa: F401

    HAVE_PYCAW = True
except Exception:
    HAVE_PYCAW = False


class PycawAudio(AudioPort):
    def _iter_target_sessions(self, exe_path_lower: str):
        if not HAVE_PYCAW:
            return []
        sessions = AudioUtilities.GetAllSessions()
        for s in sessions:
            try:
                p = s.Process
                if not p:
                    continue
                pexe = (p.exe() or "").lower()
                if pexe == exe_path_lower:
                    try:
                        vol = s._ctl.QueryInterface(ISimpleAudioVolume)
                        yield vol
                    except Exception:
                        continue
            except Exception:
                continue

    def mute(self, exe_path: str):
        exe_lower = (exe_path or "").lower()
        for vol in self._iter_target_sessions(exe_lower):
            try:
                vol.SetMute(1, None)
            except Exception:
                pass

    def unmute(self, exe_path: str):
        exe_lower = (exe_path or "").lower()
        for vol in self._iter_target_sessions(exe_lower):
            try:
                vol.SetMute(0, None)
            except Exception:
                pass
