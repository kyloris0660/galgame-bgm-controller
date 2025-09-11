# core/audio.py  — 兼容版：导出 AudioPort 接口；PycawAudio 实现 set_bulk/mute/unmute
import os
import unicodedata

try:
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

    HAVE_PYCAW = True
except Exception:
    HAVE_PYCAW = False


# ---------- 工具 ----------
def _norm_path(p: str) -> str:
    """统一大小写、分隔符，并做 Unicode 规范化（NFC）"""
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


# ---------- 对外接口（保持与 service.py 兼容） ----------
class AudioPort:
    """音频控制接口约定：Service 依赖 set_bulk(exe->bool)"""

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


# ---------- Pycaw 实现 ----------
class PycawAudio(AudioPort):
    """
    更稳的匹配策略：
      1) FULL PATH 精确匹配（规范化后）；
      2) 命中为 0 时回退到 BASENAME（仅比较文件名）。
    """

    def __init__(self):
        super().__init__()

    # 供调试脚本调用
    def _iter_target_sessions(self, target_exe: str):
        if not HAVE_PYCAW:
            return []
        sessions = AudioUtilities.GetAllSessions()
        t_norm = _norm_path(target_exe)
        t_base = _basename(target_exe)

        matched = []
        # 第一轮：full path
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

        # 第二轮：basename
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

    # 兼容：如果你在其他脚本里用过 mute_by_exe / unmute_by_exe，也能用
    def mute_by_exe(self, exe_path: str):
        self.mute(exe_path)

    def unmute_by_exe(self, exe_path: str):
        self.unmute(exe_path)

    # AudioPort 接口
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


# ---------- 可选的 Dummy 实现（供测试用，不改名保证兼容已有 tests） ----------
class DummyAudio(AudioPort):
    def __init__(self):
        self.state = {}

    def mute(self, exe_path: str):
        self.state[_norm_path(exe_path)] = True

    def unmute(self, exe_path: str):
        self.state[_norm_path(exe_path)] = False
