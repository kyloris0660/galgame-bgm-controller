import json
from types import SimpleNamespace
import pytest

from core.audio import PycawAudio, SessionBatch


class Volume:
    def __init__(self, muted=False):
        self.muted = muted
        self.writes = []
        self.offline = False

    def GetMute(self):
        return self.muted

    def SetMute(self, value, _):
        if self.offline:
            raise OSError("device invalidated")
        self.muted = bool(value)
        self.writes.append(bool(value))


def session(path, identity="s1", endpoint="speaker", muted=False):
    volume = Volume(muted)
    return SimpleNamespace(Process=SimpleNamespace(exe=lambda: path),
                           InstanceIdentifier=identity, endpoint=endpoint,
                           SimpleAudioVolume=volume, State=1)


def test_exact_path_only_and_original_mute_preserved():
    target = session(r"C:\game\Game.exe")
    other = session(r"D:\different\Game.exe", "other")
    already_muted = session(r"C:\game\Game.exe", "muted", muted=True)
    audio = PycawAudio(lambda: [target, other, already_muted])
    audio.set_bulk({r"c:/game/GAME.exe": True})
    assert target.SimpleAudioVolume.muted
    assert not other.SimpleAudioVolume.muted
    audio.set_bulk({})
    assert not target.SimpleAudioVolume.muted
    assert already_muted.SimpleAudioVolume.muted


def test_late_session_and_device_change():
    speaker = session("a.exe", "a", "speaker")
    sessions = [speaker]
    audio = PycawAudio(lambda: SessionBatch(sessions, {s.endpoint for s in sessions}))
    audio.set_bulk({"a.exe": True})
    headphones = session("a.exe", "a", "headphones")
    unrelated = session("other.exe", "unrelated", "headphones")
    sessions.extend([headphones, unrelated])
    audio.set_bulk({"a.exe": True})
    assert speaker.SimpleAudioVolume.muted and headphones.SimpleAudioVolume.muted
    assert not unrelated.SimpleAudioVolume.muted
    audio.restore_all()
    assert not headphones.SimpleAudioVolume.muted
    assert not speaker.SimpleAudioVolume.muted


def test_disconnect_pause_reconnect_restores_baseline(tmp_path):
    journal = tmp_path / "recovery.json"
    target = session("a.exe")
    batch = SessionBatch([target], {"speaker"})
    audio = PycawAudio(lambda: batch, journal)
    audio.set_bulk({"a.exe": True})
    target.SimpleAudioVolume.offline = True
    batch.sessions = []
    batch.endpoints = set()
    audio.set_bulk({})
    assert json.loads(journal.read_text())
    target.SimpleAudioVolume.offline = False
    batch.sessions = [target]
    batch.endpoints = {"speaker"}
    audio.set_bulk({})
    assert not target.SimpleAudioVolume.muted
    assert json.loads(journal.read_text()) == {}


def test_crash_recovery_uses_original_baseline(tmp_path):
    journal = tmp_path / "recovery.json"
    target = session("a.exe")
    batch = lambda: SessionBatch([target], {"speaker"})
    before_crash = PycawAudio(batch, journal)
    before_crash.set_bulk({"a.exe": True})
    restarted = PycawAudio(batch, journal)
    restarted.set_bulk({"a.exe": True})
    assert target.SimpleAudioVolume.muted
    restarted.restore_all()
    assert not target.SimpleAudioVolume.muted
    assert not json.loads(journal.read_text())


def test_partial_device_failure_keeps_working_device_controlled():
    target = session("b.exe")
    audio = PycawAudio(lambda: SessionBatch([target], {"speaker"}, ["headphone unavailable"]))
    with pytest.raises(RuntimeError, match="headphone unavailable"):
        audio.set_bulk({"b.exe": True})
    assert target.SimpleAudioVolume.muted


def test_failed_journal_write_does_not_mute(tmp_path, monkeypatch):
    target = session("a.exe")
    audio = PycawAudio(lambda: [target], tmp_path / "recovery.json")
    def fail():
        raise OSError("disk full")
    monkeypatch.setattr(audio, "_save", fail)
    with pytest.raises(RuntimeError, match="disk full"):
        audio.set_bulk({"a.exe": True})
    assert not target.SimpleAudioVolume.muted


def test_expired_session_does_not_block_exit():
    target = session("a.exe")
    audio = PycawAudio(lambda: [target])
    audio.set_bulk({"a.exe": True})
    target.State = 2
    target.SimpleAudioVolume.offline = True
    audio.restore_all()
    assert not audio._pending


def test_disconnected_device_exit_retains_recovery_record(tmp_path):
    target = session("a.exe")
    batch = SessionBatch([target], {"speaker"})
    journal = tmp_path / "recovery.json"
    audio = PycawAudio(lambda: batch, journal)
    audio.set_bulk({"a.exe": True})
    target.SimpleAudioVolume.offline = True
    batch.sessions, batch.endpoints = [], set()
    audio.restore_all()
    assert json.loads(journal.read_text())
    audio.release()
    assert not audio._owned


def test_session_disappears_before_enumeration_restores_old_handle():
    target = session("a.exe")
    batch = SessionBatch([target], {"speaker"})
    audio = PycawAudio(lambda: batch)
    audio.set_bulk({"a.exe": True})
    batch.sessions = []
    audio.set_bulk({})
    assert not target.SimpleAudioVolume.muted
    assert not audio._pending
