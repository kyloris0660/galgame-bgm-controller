import tkinter as tk
from core.windows import TaskbarApp
from ui.picker import ProcessPicker


def test_picker_multiselect_survives_filter_refresh_and_closed_window():
    root = tk.Tk()
    root.withdraw()
    apps = [TaskbarApp("a.exe", "Alpha", ()), TaskbarApp("b.exe", "Beta", ())]
    dialog = ProcessPicker(root, provider=lambda: apps)
    try:
        dialog.toggle(apps[0])
        dialog.var_q.set("Beta")
        dialog.render()
        dialog.toggle(apps[1])
        apps.pop(0)
        dialog.reload()
        assert set(dialog.selected) == {"a.exe", "b.exe"}
        dialog.on_ok()
        assert dialog.result == ["a.exe", "b.exe"]
    finally:
        if not dialog._closed:
            dialog.on_cancel()
        root.destroy()


def test_picker_existing_targets_disabled_and_cancel_does_not_add():
    root = tk.Tk()
    root.withdraw()
    dialog = ProcessPicker(root, existing=["A.exe"],
                           provider=lambda: [TaskbarApp("a.exe", "Alpha", ())])
    try:
        buttons = [c for c in dialog.grid_frame.winfo_children() if isinstance(c, tk.Button)]
        assert str(buttons[0]["state"]) == "disabled"
        assert str(dialog.btn_ok["state"]) == "disabled"
        dialog.on_cancel()
        assert dialog.result is None
    finally:
        if not dialog._closed:
            dialog.on_cancel()
        root.destroy()
