# app_gui.py
import tkinter as tk
from tkinter import ttk, messagebox
from core.config import Config
from core.audio import PycawAudio
from core.focus import WinFocus
from core.service import Service
from core.types import MuteMode
from core import debug as dbg
from ui.picker import ProcessPicker
from ui.tray import Tray


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Galgame BGM Controller – v5")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)

        # 顶部：状态 + 调试开关
        top = ttk.Frame(frm)
        top.pack(fill="x")
        self.lbl = ttk.Label(top, text="初始化…")
        self.lbl.pack(side="left")
        self.var_dbg = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top,
            text="调试日志",
            variable=self.var_dbg,
            command=lambda: dbg.set_enabled(self.var_dbg.get()),
        ).pack(side="right")

        # 监听目标列表 + 按钮
        ttk.Label(frm, text="监听目标：").pack(anchor="w", pady=(6, 0))
        self.lst = tk.Listbox(frm, height=6)
        self.lst.pack(fill="both", expand=True, pady=(2, 6))
        btns = ttk.Frame(frm)
        btns.pack(fill="x")
        ttk.Button(btns, text="选择监听软件（可多选）", command=self.pick).pack(
            side="left"
        )
        ttk.Button(btns, text="删除选中", command=self.remove_selected).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(btns, text="退出", command=self.quit).pack(side="right")

        self.cfg = Config()
        self.audio = PycawAudio()
        self.focus = WinFocus()
        # 把 on_apply 传给 Service：变化才回调，更新托盘图标
        self.svc = Service(
            self.cfg, self.audio, self.focus, interval=0.5, on_apply=self._on_apply
        )

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
            # 合并去重
            cur = {t.lower(): t for t in self.cfg.get_targets()}
            for e in dlg.result:
                cur[e.lower()] = e
            self.cfg.set_targets(list(cur.values()))
            self._refresh_targets()

    def remove_selected(self):
        idxs = list(self.lst.curselection())
        if not idxs:
            messagebox.showinfo("提示", "请选择要删除的监听进程")
            return
        targets = self.cfg.get_targets()
        # 从后往前删，避免索引位移
        for i in sorted(idxs, reverse=True):
            if 0 <= i < len(targets):
                targets.pop(i)
        self.cfg.set_targets(targets)
        self._refresh_targets()

    def _refresh_targets(self):
        self.lst.delete(0, "end")
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
            self.svc.resume()
            self.paused = False
            self.lbl.config(text=f"状态：运行中（规则：{self.cfg.get_mode().value}）")
        else:
            self.svc.pause()
            self.paused = True
            self.lbl.config(text="状态：已暂停")

    def _on_apply(self, actions: dict, muted_list: list, active_exe: str | None):
        """Service 通知有变化时调用：更新托盘图标（被静音 app 的图标）"""
        # 优先显示被静音的第一个；如果没有静音，则显示默认
        exe = muted_list[0] if muted_list else None
        self.tray.set_icon_from_exe(exe)

    def on_close(self):
        try:
            self.root.withdraw()
        except Exception:
            pass

    def quit(self):
        try:
            self.svc.stop()
        except Exception:
            pass
        try:
            self.tray.stop()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass


def main():
    App().root.mainloop()


if __name__ == "__main__":
    main()
