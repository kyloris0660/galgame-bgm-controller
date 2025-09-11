
from core.audio import DummyAudio

def test_dummy_audio_bulk():
    a = DummyAudio()
    a.set_bulk({"x.exe": True, "y.exe": False})
    assert a.state["x.exe"] is True and a.state["y.exe"] is False
