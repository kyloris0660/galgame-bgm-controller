from core.instance import SingleInstance


def test_second_launch_signals_first_and_different_config_is_independent(tmp_path):
    first = SingleInstance(str(tmp_path / "one.json"))
    second = SingleInstance(str(tmp_path / "one.json"))
    other = SingleInstance(str(tmp_path / "two.json"))
    try:
        assert first.primary
        assert not second.primary
        assert other.primary
        assert first.requested()
        assert not first.requested()
    finally:
        second.close()
        other.close()
        first.close()
