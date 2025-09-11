import json
import os
from typing import List
from .types import MuteMode

DEFAULT_PATH = os.path.join(os.path.expanduser("~"), "gal_audio_controller_config.json")


class Config:
    def __init__(self, path: str = DEFAULT_PATH):
        self.path = path
        if not os.path.exists(self.path):
            self.write(
                {"selected_procs": [], "mute_mode": MuteMode.NOT_FOREGROUND.value}
            )

    def read(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"selected_procs": [], "mute_mode": MuteMode.NOT_FOREGROUND.value}

    def write(self, data):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def get_targets(self) -> List[str]:
        return list(self.read().get("selected_procs", []))

    def set_targets(self, exe_list: List[str]):
        exe_list = [e for e in exe_list if isinstance(e, str) and e.strip()]
        seen = {}
        for e in exe_list:
            seen[e] = True
        data = self.read()
        data["selected_procs"] = list(seen.keys())
        self.write(data)

    def get_mode(self) -> MuteMode:
        val = self.read().get("mute_mode", MuteMode.NOT_FOREGROUND.value)
        try:
            return MuteMode(val)
        except Exception:
            return MuteMode.NOT_FOREGROUND

    def set_mode(self, mode: MuteMode):
        data = self.read()
        data["mute_mode"] = mode.value
        self.write(data)
