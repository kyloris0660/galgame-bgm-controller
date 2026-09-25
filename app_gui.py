import argparse
import ctypes
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import queue
import sys
import tkinter as tk
from tkinter import ttk, messagebox

from PIL import ImageTk, Image
from core.config import Config, DEFAULT_PATH
from core.audio import PycawAudio
from core.focus import WinFocus
from core.icons import get_exe_icon_pil
from core.instance import SingleInstance
from core.paths import exe_key
from core.service import Service
from core.types import MuteMode
from ui.picker import ProcessPicker
from ui.tray import Tray, state_label
from ui.theme import configure, MODE_LABELS, FONT

VERSION = "2.0.0"


class App:
    def __init__(self, cfg=None, instance=None):
        self.cfg = cfg or Config()
        self.instance = instance
        self.root = tk.Tk()
        self.root.title(f"Galgame BGM Controller · {VERSION}")
        self.root.geometry("1120x800")
        self.root.minsize(900, 680)
        configure(self.root)
        asset = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "assets" / "app.ico"
        if asset.exists():
            self.root.iconbitmap(str(asset))
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.root.report_callback_exception = self._callback_error
        self.events = queue.SimpleQueue()
        self.states = {}
        self._photos = {}
        self._rows = {}
        self._picker = None
        self._closing = False
        self._error = None

        header = ttk.Frame(self.root, padding=(28, 24, 28, 18))
        header.pack(fill="x")
        ttk.Label(header, text="让音乐跟随游戏", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="启动一次，每个游戏独立控制。关闭窗口后继续在托盘运行。",
                  style="Muted.TLabel").pack(anchor="w", pady=(6, 15))
        bar = ttk.Frame(header)
        bar.pack(fill="x")
        ttk.Button(bar, text="＋ 添加游戏", style="Accent.TButton", command=self.pick).pack(side="left")
        self.pause_button = ttk.Button(bar, text="暂停全部", command=self.toggle_pause)
        self.pause_button.pack(side="left", padx=10)
        self.default_mode = tk.StringVar(value=MODE_LABELS[self.cfg.get_mode().value])
        mode_box = ttk.Combobox(bar, textvariable=self.default_mode, values=list(MODE_LABELS.values()),
                               state="readonly", width=20)
        mode_box.pack(side="right")
        mode_box.bind("<<ComboboxSelected>>", lambda _: self.change_mode(
            next(k for k, v in MODE_LABELS.items() if v == self.default_mode.get())))
        ttk.Label(bar, text="默认规则  ", style="Muted.TLabel").pack(side="right")

        library = ttk.Frame(self.root, padding=(28, 0, 28, 0))
        library.pack(fill="both", expand=True)
        self.counter = ttk.Label(library, text="游戏库", font=(FONT, 12, "bold"))
        self.counter.pack(anchor="w", pady=(0, 8))
        table = ttk.Frame(library)
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, columns=("status", "mode"), selectmode="browse", height=4)
        self.tree.heading("#0", text="游戏")
        self.tree.heading("status", text="当前状态")
        self.tree.heading("mode", text="静音规则")
        self.tree.column("#0", width=400, minwidth=200)
        self.tree.column("status", width=110, minwidth=90, stretch=False)
        self.tree.column("mode", width=190, minwidth=140, stretch=False)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<<TreeviewSelect>>", self._selection_changed)

        controls = ttk.Frame(self.root, padding=(28, 14, 28, 10))
        controls.pack(fill="x")
        self.target_pause = ttk.Button(controls, text="暂停此游戏", command=self.toggle_selected)
        self.target_pause.pack(side="left")
        self.target_mode = tk.StringVar(value="跟随默认规则")
        self.target_box = ttk.Combobox(controls, textvariable=self.target_mode, state="readonly",
                                      values=["跟随默认规则", *MODE_LABELS.values()], width=20)
        self.target_box.pack(side="left", padx=10)
        self.target_box.bind("<<ComboboxSelected>>", self._selected_mode)
        self.remove_button = ttk.Button(controls, text="移出游戏库", command=self.remove_selected)
        self.remove_button.pack(side="left")
        ttk.Button(controls, text="退出控制器", command=self.quit).pack(side="right")

        footer = ttk.Frame(self.root, padding=(28, 6, 28, 20))
        footer.pack(fill="x")
        self.status = ttk.Label(footer, text="正在启动…", style="Muted.TLabel", wraplength=800)
        self.status.pack(anchor="w")
        ttk.Label(footer, text="未启动的游戏会保留在库中；再次启动时，托盘图标会自动出现。",
                  style="Muted.TLabel").pack(anchor="w", pady=(5, 0))

        self.svc = Service(self.cfg, PycawAudio(journal_path=Path(self.cfg.path).with_suffix(".audio-state.json")), WinFocus(),
                           on_state=lambda states: self.events.put(("state", states)),
                           on_error=lambda error: self.events.put(("error", error)))
        # No Tcl/Tk calls from the service or pystray threads.
        def dispatch(callback):
            return lambda: self.events.put(("call", callback))
        self.tray = Tray(
            on_select=dispatch(self.pick), on_pause=dispatch(self.toggle_pause),
            on_quit=dispatch(self.quit), on_show=dispatch(self.show_main),
            on_target_pause=lambda exe: self.events.put(("call", lambda: self.svc.toggle_target(exe))),
            on_target_stop=lambda exe: self.events.put(("call", lambda: self.svc.dismiss_target(exe))),
            on_mode_change=lambda mode, exe: self.events.put(("call", lambda: self.change_mode(mode, exe))),
            get_name=self.cfg.get_name)
        self._refresh_targets()
        self.root.after(0, self.start)

    def _callback_error(self, kind, value, trace):
        logging.getLogger(__name__).error("UI callback failed", exc_info=(kind, value, trace))
        messagebox.showerror("操作未完成", str(value), parent=self.root)

    def start(self):
        try:
            self.tray.start()
        except Exception as exc:
            self._error = f"托盘启动失败：{exc}"
            self._update_status()
            return
        self.svc.start()
        self._drain()

    def _drain(self):
        if self._closing:
            return
        try:
            if self.instance and self.instance.requested():
                self.show_main()
            while not self.events.empty():
                kind, value = self.events.get()
                if kind == "call":
                    value()
                    if self._closing:
                        return
                elif kind == "state":
                    self.states = value
                    self.tray.sync(value, self.svc.paused)
                    self._refresh_targets()
                else:
                    self._error = value
            self._update_status()
        finally:
            if not self._closing:
                self.root.after(100, self._drain)

    def _update_status(self):
        self.pause_button.config(text="继续全部" if self.svc.paused else "暂停全部")
        self.status.config(text=self._error or (
            "全部已暂停 · 不再自动静音" if self.svc.paused else
            f"正在监听 {sum(s.controlled for s in self.states.values())} 个游戏 · 音频设备和新游戏会自动重新检测"))

    def hide_to_tray(self):
        if self.tray.icon:
            self.root.withdraw()

    def show_main(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        if self._picker and self._picker.winfo_exists():
            self._picker.lift()

    def pick(self):
        if self._picker and self._picker.winfo_exists():
            self._picker.lift()
            return
        was_hidden = self.root.state() == "withdrawn"
        self.show_main()
        dialog = ProcessPicker(self.root, existing=self.cfg.get_targets())
        self._picker = dialog
        self.root.wait_window(dialog)
        self._picker = None
        if self._closing:
            return
        if dialog.result:
            current = {exe_key(t): t for t in self.cfg.get_targets()}
            current.update({exe_key(t): t for t in dialog.result})
            self.cfg.set_targets(list(current.values()))
            for exe in dialog.result:
                self.cfg.set_name(exe, dialog.names[exe_key(exe)])
            self.svc.wake()
            self._refresh_targets()
        if was_hidden:
            self.hide_to_tray()

    def selected(self):
        selection = self.tree.selection()
        return self._rows.get(selection[0]) if selection else None

    def _refresh_targets(self):
        selected = self.selected()
        scroll_position = self.tree.yview()[0]
        self.tree.delete(*self.tree.get_children())
        self._rows.clear()
        for index, exe in enumerate(self.cfg.get_targets()):
            key = exe_key(exe)
            state = self.states.get(key)
            if key not in self._photos:
                self._photos[key] = ImageTk.PhotoImage(get_exe_icon_pil(exe).resize((36, 36), Image.Resampling.LANCZOS))
            row = str(index)
            self._rows[row] = exe
            self.tree.insert("", "end", iid=row, text="  " + self.cfg.get_name(exe), image=self._photos[key],
                             values=(state_label(state) if state else "未启动", MODE_LABELS[self.cfg.get_mode(exe).value]))
            if selected and exe_key(selected) == key:
                self.tree.selection_set(row)
        count = len(self._rows)
        self.counter.config(text=f"游戏库  ·  {count} 个已添加  /  {len(self.states)} 个运行中" if count else "游戏库为空 · 点击「添加游戏」开始")
        self.tree.yview_moveto(scroll_position)
        self._selection_changed()

    def _selection_changed(self, *_):
        exe = self.selected()
        self.remove_button.config(state="normal" if exe else "disabled")
        self.target_box.config(state="readonly" if exe else "disabled")
        state = self.states.get(exe_key(exe)) if exe else None
        self.target_pause.config(state="normal" if state and not self.svc.paused else "disabled",
                                 text="继续此游戏" if state and state.paused else "暂停此游戏")
        if exe:
            override = self.cfg.read().get("apps", {}).get(exe_key(exe), {}).get("mode")
            self.target_mode.set(MODE_LABELS.get(override, "跟随默认规则"))

    def _selected_mode(self, *_):
        exe = self.selected()
        if exe:
            mode = next((k for k, v in MODE_LABELS.items() if v == self.target_mode.get()), None)
            self.change_mode(mode, exe)

    def remove_selected(self):
        exe = self.selected()
        if exe:
            self.cfg.set_targets([t for t in self.cfg.get_targets() if exe_key(t) != exe_key(exe)])
            self.svc.wake()
            self._refresh_targets()

    def toggle_selected(self):
        exe = self.selected()
        if exe:
            self.svc.toggle_target(exe)

    def change_mode(self, mode, exe=None):
        self.cfg.set_mode(MuteMode(mode) if mode else None, exe)
        self.svc.wake()
        self._refresh_targets()

    def toggle_pause(self):
        self.svc.resume() if self.svc.paused else self.svc.pause()
        self.tray.sync(self.states, self.svc.paused)
        self._selection_changed()
        self._update_status()

    def quit(self):
        if self._closing:
            return
        self.status.config(text="正在恢复音频并退出…")
        self.root.update_idletasks()
        try:
            self.svc.stop()
        except Exception as exc:
            messagebox.showerror("音频恢复未完成", str(exc), parent=self.root)
            return
        self._closing = True
        self.tray.stop()
        if self._picker:
            self._picker.on_cancel()
        self.root.destroy()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=DEFAULT_PATH, help="Use a separate configuration file")
    parser.add_argument("--smoke-test", type=float, metavar="SECONDS", help="Start normally, then exit gracefully")
    parser.add_argument("--diagnostics", metavar="JSON", help="Write read-only packaging diagnostics and exit")
    args = parser.parse_args()
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except OSError:
        pass
    if args.diagnostics:
        from core.windows import taskbar_apps
        from core.audio import scan_sessions
        scan = scan_sessions()
        data = {"version": VERSION, "executable": sys.executable,
                "taskbar_apps": len(taskbar_apps()), "audio_sessions": len(scan.sessions),
                "audio_devices": len(scan.endpoints), "audio_errors": scan.errors}
        Path(args.diagnostics).write_text(json.dumps(data, indent=2), encoding="utf-8")
        return
    instance = SingleInstance(args.config)
    if not instance.primary:
        instance.close()
        return
    try:
        log_path = Path(args.config).with_suffix(".log")
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handler = RotatingFileHandler(log_path, maxBytes=1024 * 1024, backupCount=2, encoding="utf-8")
        logging.basicConfig(level=logging.INFO, handlers=[handler],
                            format="%(asctime)s %(levelname)s %(name)s: %(message)s")
        cfg = Config(args.config)
        app = App(cfg, instance)
        if args.smoke_test is not None:
            app.root.after(max(1, int(args.smoke_test * 1000)), app.quit)
        app.root.mainloop()
    except Exception as exc:
        logging.exception("Startup failed")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Galgame BGM 启动失败", f"{exc}\n\n配置文件：{args.config}", parent=root)
        root.destroy()
        raise
    finally:
        instance.close()


if __name__ == "__main__":
    main()
