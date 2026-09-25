"""Windows executable identity shared by configuration, windows and audio."""
import ntpath


def exe_key(path: str) -> str:
    return ntpath.normcase(ntpath.normpath(path.strip())) if path.strip() else ""


def app_name(path: str) -> str:
    return ntpath.splitext(ntpath.basename(path))[0]
