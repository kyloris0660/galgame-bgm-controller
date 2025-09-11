from dataclasses import dataclass
from enum import Enum
from typing import List, Optional


class MuteMode(str, Enum):
    NOT_FOREGROUND = "not_foreground"
    MINIMIZED_ONLY = "minimized_only"


@dataclass
class TargetStatus:
    exe: str
    is_foreground: bool
    minimized: bool


@dataclass
class EventSnapshot:
    active_exe: Optional[str]
    targets: List[TargetStatus]
