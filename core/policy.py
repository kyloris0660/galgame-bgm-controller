from .types import EventSnapshot, MuteMode


def should_mute(mode: MuteMode, is_foreground: bool, minimized: bool) -> bool:
    if mode == MuteMode.MINIMIZED_ONLY:
        return minimized
    return not is_foreground  # NOT_FOREGROUND


def apply_policy(snapshot: EventSnapshot, mode: MuteMode):
    """Return dict exe->mute(bool)"""
    actions = {}
    for t in snapshot.targets:
        actions[t.exe] = should_mute(mode, t.is_foreground, t.minimized)
    return actions
