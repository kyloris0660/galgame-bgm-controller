
import tkinter as tk
from tkinter import ttk
from core.config import Config
from core.audio import PycawAudio
from core.focus import WinFocus
from core.service import Service
from core.types import MuteMode
from ui.picker import ProcessPicker
from ui.tray import Tray

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Galgame BGM Controller – v5")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        frm = ttk.Frame(self.root, padding=10); frm.pack(fill="both", expand=True)
        self.lbl = ttk.Label(frm, text="初始化…"); self.lbl.pack(anchor="w")
        self.lst = tk.Listbox(frm, height=5); self.lst.pack(fill="both", expand=True, pady=(6,6))
        btns = ttk.Frame(frm); btns.pack(fill="x")
        ttk.Button(btns, text="选择监听软件（可多选）", command=self.pick).pack(side="left")
        ttk.Button(btns, text="退出", command=self.quit).pack(side="right")

        self.cfg = Config()
        self.svc = Service(self.cfg, PycawAudio(), WinFocus(), interval=0.5)
        self.tray = Tray(
            on_select=self.pick,
            on_pause=self.toggle_pause,
            on_quit=self.quit,
            on_mode_change=self.change_mode,
            get_mode=lambda: self.cfg.get_mode().value,
        )
        self.paused = False
        self.root.after(0, self.start)

    def start(self):
        self.tray.start()
        self._refresh_targets()
        self.lbl.config(text=f"状态：运行中（规则：{self.cfg.get_mode().value}）")
        self.svc.start()

    def pick(self):
        dlg = ProcessPicker(self.root, multiselect=True)
        self.root.wait_window(dlg)
        if dlg.result:
            self.cfg.set_targets(dlg.result)
            self._refresh_targets()

    def _refresh_targets(self):
        self.lst.delete(0,"end")
        for t in self.cfg.get_targets():
            self.lst.insert("end", t)

    def change_mode(self, mode_str: str):
        try:
            self.cfg.set_mode(MuteMode(mode_str))
        except Exception:
            self.cfg.set_mode(MuteMode.NOT_FOREGROUND)
        self.lbl.config(text=f"状态：运行中（规则：{self.cfg.get_mode().value}）")

    def toggle_pause(self):
        if self.paused:
            self.svc.resume(); self.paused=False
            self.lbl.config(text=f"状态：运行中（规则：{self.cfg.get_mode().value}）")
        else:
            self.svc.pause(); self.paused=True
            self.lbl.config(text="状态：已暂停")

    def on_close(self):
        # 隐藏到托盘
        try: self.root.withdraw()
        except Exception: pass

    def quit(self):
        try: self.svc.stop()
        except Exception: pass
        try: self.tray.stop()
        except Exception: pass
        try: self.root.destroy()
        except Exception: pass

def main():
    App().root.mainloop()

if __name__ == "__main__":
    main()
