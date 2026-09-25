"""One controller per configuration; subsequent launches wake its library window."""
import hashlib
import os

import win32api
import win32event
import winerror


class SingleInstance:
    def __init__(self, config_path):
        identity = hashlib.sha256(os.path.normcase(os.path.abspath(config_path)).encode()).hexdigest()[:24]
        self.mutex = win32event.CreateMutex(None, False, "Local\\GalBGM-" + identity)
        self.primary = win32api.GetLastError() != winerror.ERROR_ALREADY_EXISTS
        self.event = win32event.CreateEvent(None, False, False, "Local\\GalBGM-Show-" + identity)
        if not self.primary:
            win32event.SetEvent(self.event)

    def requested(self):
        return win32event.WaitForSingleObject(self.event, 0) == win32event.WAIT_OBJECT_0

    def close(self):
        self.event.Close()
        self.mutex.Close()
