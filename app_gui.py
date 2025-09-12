import sys
import tkinter as tk
from tkinter import ttk, messagebox
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
        # X 号不退出，只最小化到托盘
        self.root.protocol("WM_DELETE_WINDOW", self.on_close_to_tray)

        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)

        top = ttk.Frame(frm)
        top.pack(fill="x")
        self.lbl = ttk.Label(top, text="初始化…")
        self.lbl.pack(side="left")

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
        self.svc = Service(
            self.cfg, self.audio, self.focus, interval=0.5, on_apply=self._on_apply
        )

        # 关键：托盘回调一律切回 Tk 主线程
        self.tray = Tray(
            on_select=lambda: self.root.after(0, self.pick),
            on_pause=lambda: self.root.after(0, self.toggle_pause),
            on_quit=lambda: self.root.after(0, self.quit),
            on_mode_change=lambda s: self.root.after(0, lambda: self.change_mode(s)),
            get_mode=lambda: self.cfg.get_mode().value,
            on_show=lambda: self.root.after(0, self.show_main),  # 传入新回调
        )

        # 状态
        self.paused = False
        self._picker = None  # 选择器单例
        self._hidden_to_tray = False  # 是否被托盘隐藏

        self.root.after(0, self.start)

    # ---------- 托盘最小化/恢复 ----------
    def on_close_to_tray(self):
        self.hide_to_tray()

    def hide_to_tray(self):
        try:
            self._hidden_to_tray = True
            # withdraw: 从任务栏消失，仅托盘驻留
            self.root.withdraw()
        except Exception:
            pass

    def show_main(self):
        try:
            self._hidden_to_tray = False
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            pass

    def _ensure_for_dialog(self):
        """
        确保能够创建模态对话框：
        - 若窗体 withdraw/iconify，则临时显示出来
        - 返回一个回调，供对话框关闭后恢复原状态
        """
        was_hidden = False
        try:
            state = str(self.root.state())
            if state in ("withdrawn", "iconic"):
                was_hidden = True
        except Exception:
            was_hidden = self._hidden_to_tray

        if was_hidden:
            self.show_main()
            self.root.update_idletasks()

        def restore():
            if was_hidden or self._hidden_to_tray:
                self.hide_to_tray()

        return restore

    # ---------- 生命周期 ----------
    def start(self):
        self.tray.start()
        self._refresh_targets()
        self.lbl.config(text=f"状态：运行中（规则：{self.cfg.get_mode().value}）")
        self.svc.start()

    # ---------- 进程选择 ----------
    def pick(self):
        # 如果已经打开，就唤醒置顶
        try:
            if self._picker and self._picker.winfo_exists():
                try:
                    self._picker.deiconify()
                    self._picker.lift()
                    self._picker.focus_force()
                    self._picker.attributes("-topmost", True)
                    self._picker.after(
                        200, lambda: self._picker.attributes("-topmost", False)
                    )
                except Exception:
                    pass
                return
        except Exception:
            pass

        # 确保可以创建模态对话框
        restore = self._ensure_for_dialog()

        # 新建对话框
        try:
            dlg = ProcessPicker(self.root, multiselect=True)
            self._picker = dlg
        except Exception:
            messagebox.showerror("错误", "无法打开进程选择器，请重试。")
            self._picker = None
            restore()
            return

        # 等待对话框关闭
        self.root.wait_window(dlg)
        self._picker = None
        restore()

        if getattr(dlg, "result", None):
            cur = {t.lower(): t for t in self.cfg.get_targets()}
            for e in dlg.result:
                cur[e.lower()] = e
            self.cfg.set_targets(list(cur.values()))
            self._refresh_targets()

    # ---------- 列表维护 ----------
    def remove_selected(self):
        idxs = list(self.lst.curselection())
        if not idxs:
            messagebox.showinfo("提示", "请选择要删除的监听进程")
            return
        targets = self.cfg.get_targets()
        for i in sorted(idxs, reverse=True):
            if 0 <= i < len(targets):
                targets.pop(i)
        self.cfg.set_targets(targets)
        self._refresh_targets()

    def _refresh_targets(self):
        self.lst.delete(0, "end")
        for t in self.cfg.get_targets():
            self.lst.insert("end", t)

    # ---------- 行为 ----------
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
        exe = muted_list[0] if muted_list else None
        # 后台 -> 主线程更新托盘图标
        self.root.after(0, lambda: self.tray.set_icon_from_exe(exe if exe else None))

    # ---------- 退出 ----------
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
        try:
            sys.exit(0)
        except Exception:
            import os

            os._exit(0)


def main():
    App().root.mainloop()


if __name__ == "__main__":
    main()
