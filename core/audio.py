import os
import unicodedata

try:
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

    HAVE_PYCAW = True
except Exception:
    HAVE_PYCAW = False


def _norm_path(p: str) -> str:
    if not p:
        return ""
    try:
        p = os.path.normpath(p)
        p = unicodedata.normalize("NFC", p)
        p = p.replace("/", "\\")
        return p.lower()
    except Exception:
        return (p or "").lower()


def _basename(p: str) -> str:
    try:
        return os.path.basename(p).lower()
    except Exception:
        return (p or "").lower()


class AudioPort:
    def set_bulk(self, mutes: dict):
        for exe, flag in (mutes or {}).items():
            if flag:
                self.mute(exe)
            else:
                self.unmute(exe)

    def mute(self, exe_path: str):
        raise NotImplementedError

    def unmute(self, exe_path: str):
        raise NotImplementedError


class PycawAudio(AudioPort):
    def __init__(self):
        super().__init__()

    def _iter_target_sessions(self, target_exe: str):
        if not HAVE_PYCAW:
            return []
        sessions = AudioUtilities.GetAllSessions()
        t_norm = _norm_path(target_exe)
        t_base = _basename(target_exe)
        matched = []
        for s in sessions:
            try:
                p = s.Process
                if not p:
                    continue
                pexe_norm = _norm_path(p.exe() or "")
                if pexe_norm == t_norm:
                    vol = s._ctl.QueryInterface(ISimpleAudioVolume)
                    matched.append(vol)
            except Exception:
                continue
        if matched:
            return matched
        for s in sessions:
            try:
                p = s.Process
                if not p:
                    continue
                pexe_base = _basename(p.exe() or "")
                if pexe_base == t_base:
                    vol = s._ctl.QueryInterface(ISimpleAudioVolume)
                    matched.append(vol)
            except Exception:
                continue
        return matched

    def mute_by_exe(self, exe_path: str):
        self.mute(exe_path)

    def unmute_by_exe(self, exe_path: str):
        self.unmute(exe_path)

    def mute(self, exe_path: str):
        if not HAVE_PYCAW:
            return
        for vol in self._iter_target_sessions(exe_path):
            try:
                vol.SetMute(1, None)
            except Exception:
                pass

    def unmute(self, exe_path: str):
        if not HAVE_PYCAW:
            return
        for vol in self._iter_target_sessions(exe_path):
            try:
                vol.SetMute(0, None)
            except Exception:
                pass

    def unmute_multi(self, exe_list):
        if not HAVE_PYCAW:
            return
        for exe in exe_list or []:
            for vol in self._iter_target_sessions(exe):
                try:
                    vol.SetMute(0, None)
                except Exception:
                    pass
