import copy
import json
import os
from pathlib import Path
import tempfile
import threading

from .paths import exe_key, app_name
from .types import MuteMode

DEFAULT_PATH = os.path.join(os.path.expanduser("~"), "gal_audio_controller_config.json")


class Config:
    """In-memory settings with atomic writes; legacy selected_procs remains valid."""

    def __init__(self, path: str = DEFAULT_PATH):
        self.path = str(Path(path).resolve())
        self._lock = threading.RLock()
        if Path(self.path).exists():
            # Never silently replace a damaged user configuration with an empty one.
            with open(self.path, encoding="utf-8-sig") as stream:
                data = json.load(stream)
            if not isinstance(data, dict) or not isinstance(data.get("selected_procs", []), list):
                raise ValueError("配置格式无效；请备份并检查配置文件。")
        else:
            data = {"selected_procs": [], "mute_mode": MuteMode.NOT_FOREGROUND.value}
        self._data = data

    def read(self):
        with self._lock:
            return copy.deepcopy(self._data)

    def write(self, data):
        with self._lock:
            parent = Path(self.path).parent
            parent.mkdir(parents=True, exist_ok=True)
            fd, temporary = tempfile.mkstemp(prefix=".bgm-", suffix=".json", dir=parent)
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as stream:
                    json.dump(data, stream, ensure_ascii=False, indent=2)
                    stream.flush()
                    os.fsync(stream.fileno())
                os.replace(temporary, self.path)
                self._data = copy.deepcopy(data)
            finally:
                if os.path.exists(temporary):
                    os.unlink(temporary)

    def get_targets(self):
        with self._lock:
            return list({exe_key(e): e for e in self._data.get("selected_procs", [])
                         if isinstance(e, str) and e.strip()}.values())

    def set_targets(self, exe_list):
        with self._lock:
            data = self.read()
            targets = {exe_key(e): e.strip() for e in exe_list if isinstance(e, str) and e.strip()}
            data["selected_procs"] = list(targets.values())
            data["apps"] = {k: v for k, v in data.get("apps", {}).items() if k in targets}
            self.write(data)

    def get_mode(self, exe=None):
        with self._lock:
            value = self._data.get("mute_mode", MuteMode.NOT_FOREGROUND.value)
            if exe:
                value = self._data.get("apps", {}).get(exe_key(exe), {}).get("mode") or value
            try:
                return MuteMode(value)
            except ValueError:
                return MuteMode.NOT_FOREGROUND

    def set_mode(self, mode, exe=None):
        with self._lock:
            data = self.read()
            if exe:
                data.setdefault("apps", {}).setdefault(exe_key(exe), {})["mode"] = mode.value if mode else None
            else:
                data["mute_mode"] = MuteMode(mode).value
            self.write(data)

    def get_name(self, exe):
        with self._lock:
            return self._data.get("apps", {}).get(exe_key(exe), {}).get("name") or app_name(exe)

    def set_name(self, exe, name):
        with self._lock:
            data = self.read()
            data.setdefault("apps", {}).setdefault(exe_key(exe), {})["name"] = name
            self.write(data)
