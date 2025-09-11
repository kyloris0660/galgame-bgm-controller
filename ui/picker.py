import os
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import psutil

    HAVE = True
except Exception:
    HAVE = False


class ProcessPicker(tk.Toplevel):
    def __init__(self, master, multiselect=True):
        super().__init__(master)
        self.title("选择要监听的进程")
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.geometry("560x420")

        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="搜索：").pack(side="left")
        self.var_q = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.var_q)
        ent.pack(side="left", fill="x", expand=True, padx=(4, 8))
        ent.bind("<KeyRelease>", self._refresh)

        mid = ttk.Frame(self, padding=(8, 0, 8, 8))
        mid.pack(fill="both", expand=True)
        selectmode = "extended" if multiselect else "browse"
        self.listbox = tk.Listbox(mid, selectmode=selectmode)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=sb.set)

        bot = ttk.Frame(self, padding=8)
        bot.pack(fill="x")
        self.btn_ok = ttk.Button(bot, text="确定", command=self.on_ok)
        self.btn_ok.pack(side="right")
        ttk.Button(bot, text="取消", command=self.on_cancel).pack(
            side="right", padx=(0, 6)
        )

        self._bind_validation()
        self._all = []
        self._load()
        self._refresh()

        self.transient(master)
        self.grab_set()
        ent.focus_set()

    def _bind_validation(self):
        def refresh(*_):
            self.btn_ok.config(
                state=("normal" if self.listbox.curselection() else "disabled")
            )

        self.listbox.bind("<<ListboxSelect>>", refresh)
        self.listbox.bind("<KeyRelease>", refresh)
        self.listbox.bind("<ButtonRelease-1>", refresh)
        self._refresh_ok = refresh
        refresh()

    def _load(self):
        self._all.clear()
        if not HAVE:
            return
        seen = set()
        for p in psutil.process_iter(["name", "exe"]):
            exe = (p.info.get("exe") or "").strip()
            name = (p.info.get("name") or "").strip() or os.path.basename(exe)
            if not exe:
                continue
            key = exe.lower()
            if key in seen:
                continue
            seen.add(key)
            self._all.append((name, exe))
        self._all.sort(key=lambda x: x[0].lower())

    def _refresh(self, *_):
        q = (self.var_q.get() or "").lower()
        self.listbox.delete(0, "end")
        for name, exe in self._all:
            if q in name.lower() or q in exe.lower():
                self.listbox.insert("end", f"{name}  —  {exe}")
        self._refresh_ok()

    def on_ok(self):
        idxs = self.listbox.curselection()
        if not idxs:
            messagebox.showinfo("提示", "请先选择至少一个进程")
            return
        res = []
        for i in idxs:
            t = self.listbox.get(i)
            exe = t.split("—", 1)[1].strip() if "—" in t else t.strip()
            res.append(exe)
        self.result = res
        self.grab_release()
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.grab_release()
        self.destroy()
