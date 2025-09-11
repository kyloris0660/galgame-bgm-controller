# core/debug.py
import json
import threading
import time

_enabled = False
_last = None
_lock = threading.Lock()


def set_enabled(flag: bool):
    global _enabled
    _enabled = bool(flag)


def log_change(tag: str, payload: dict):
    """只有与上次不同才打印；打印带时间戳与 tag。"""
    global _last
    if not _enabled:
        return
    with _lock:
        s = json.dumps(
            {"tag": tag, "payload": payload}, ensure_ascii=False, sort_keys=True
        )
        if s != _last:
            _last = s
            print(time.strftime("[%H:%M:%S]"), tag, s)
