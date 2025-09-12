# ui/picker.py
import os
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import ImageTk
from core.icons import get_exe_icon_pil

try:
    import psutil
    import win32gui
    import win32process

    HAVE = True
except Exception:
    HAVE = False


class ProcessPicker(tk.Toplevel):
    """带图标的进程选择器：前台优先淡黄高亮，右侧预览大图标+路径，支持多选。"""

    def __init__(self, master, multiselect=True):
        super().__init__(master)
        self.title("选择要监听的进程（带图标）")
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.geometry("820x520")

        self._photos = {}
        self._items = []
        self._refresh_job = None
        self._closed = False

        # 顶栏：搜索 + “优先显示推荐”
        bar = ttk.Frame(self, padding=(10, 8, 10, 6))
        bar.pack(fill="x")
        ttk.Label(bar, text="搜索：", font=("", 10, "bold")).pack(side="left")
        self.var_q = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.var_q, width=40)
        ent.pack(side="left", fill="x", expand=True, padx=(4, 8))
        # 输入防抖：减少频繁刷新导致的闪烁
        self.var_q.trace_add("write", lambda *_: self._schedule_refresh(120))
        self.var_hot = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            bar,
            text="优先显示推荐",
            variable=self.var_hot,
            command=lambda: self._schedule_refresh(0),
        ).pack(side="left")

        # 主区：左树 + 右预览
        main = ttk.Frame(self, padding=(10, 0, 10, 10))
        main.pack(fill="both", expand=True)
        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)
        columns = ("name", "exe")
        self.tree = ttk.Treeview(
            left,
            columns=columns,
            show="tree headings",
            selectmode=("extended" if multiselect else "browse"),
        )
        self.tree.heading("#0", text="图标")
        self.tree.heading("name", text="名称")
        self.tree.heading("exe", text="可执行文件路径")
        self.tree.column("#0", width=48, stretch=False)
        self.tree.column("name", width=220)
        self.tree.column("exe", width=420)
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.tag_configure("hot", background="#fffce6")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda e: self.on_ok())

        right = ttk.Frame(main, width=260)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)
        self.lbl_big = ttk.Label(right)
        self.lbl_big.pack(pady=(4, 8))
        self.lbl_title = ttk.Label(right, text="未选择", font=("", 12, "bold"))
        self.lbl_title.pack(anchor="w")
        self.lbl_exe = ttk.Label(right, text="", wraplength=240)
        self.lbl_exe.pack(anchor="w", pady=(2, 10))

        # 底部按钮
        bot = ttk.Frame(self, padding=(10, 0, 10, 10))
        bot.pack(fill="x")
        self.btn_ok = ttk.Button(bot, text="确定", command=self.on_ok)
        self.btn_ok.pack(side="right")
        ttk.Button(bot, text="取消", command=self.on_cancel).pack(
            side="right", padx=(0, 6)
        )

        # 初次加载与刷新
        self._load()
        self._schedule_refresh(0)

        # 模态
        self.transient(master)
        self.grab_set()
        ent.focus_set()

    def _get_foreground_exe(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return None
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            import psutil as _ps

            return (_ps.Process(pid).exe() or "").lower()
        except Exception:
            return None

    def _load(self):
        self._items.clear()
        if not HAVE:
            return
        seen = set()
        fg = self._get_foreground_exe()
        try:
            iterator = psutil.process_iter(["name", "exe"])
        except Exception:
            iterator = []
        for p in iterator:
            try:
                exe = (p.info.get("exe") or "").strip()
                if not exe:
                    continue
                key = exe.lower()
                if key in seen:
                    continue
                seen.add(key)
                name = (p.info.get("name") or os.path.basename(exe)) or exe
                score = 100 if key == fg else 0
                self._items.append({"name": name, "exe": exe, "score": score})
            except Exception:
                continue
        for it in sorted(self._items, key=lambda x: -x["score"])[:12]:
            try:
                self._get_photo(it["exe"], 20)
            except Exception:
                pass

    def _get_photo(self, exe, size):
        key = (exe, size)
        if key in self._photos:
            return self._photos[key]
        pil = get_exe_icon_pil(exe, large=True)
        if pil.size != (size, size):
            pil = pil.resize((size, size), resample=1)
        ph = ImageTk.PhotoImage(pil)
        self._photos[key] = ph
        return ph

    def _schedule_refresh(self, delay_ms):
        if self._refresh_job:
            try:
                self.after_cancel(self._refresh_job)
            except Exception:
                pass
            self._refresh_job = None
        if self._closed or not self.winfo_exists():
            return
        self._refresh_job = self.after(delay_ms, self._refresh)

    def _refresh(self, *_):
        if self._closed or not self.winfo_exists():
            return
        q = (self.var_q.get() or "").lower()
        only_hot = self.var_hot.get()
        items = [
            it
            for it in self._items
            if (not q or q in it["name"].lower() or q in it["exe"].lower())
        ]
        items.sort(key=lambda x: (-x["score"], x["name"].lower(), x["exe"].lower()))

        # 关键：先清空再插入，避免“越刷越多”
        try:
            self.tree.delete(*self.tree.get_children())
        except Exception:
            return

        for it in items:
            tags = ("hot",) if it["score"] >= 100 else ()
            img = None
            try:
                img = self._get_photo(it["exe"], 18)
            except Exception:
                img = None
            if only_hot and not tags:
                iid = self.tree.insert(
                    "", "end", text="", image=img, values=(it["name"], it["exe"])
                )
                self.tree.item(iid, tags=("dim",))
                self.tree.tag_configure("dim", foreground="#666666")
            else:
                self.tree.insert(
                    "",
                    "end",
                    text="",
                    image=img,
                    values=(it["name"], it["exe"]),
                    tags=tags,
                )

        self._refresh_ok()

    def _refresh_ok(self, *_):
        has_sel = bool(self.tree.selection())
        try:
            self.btn_ok.configure(state=("normal" if has_sel else "disabled"))
        except Exception:
            pass

    def _on_select(self, *_):
        self._refresh_ok()
        sels = self.tree.selection()
        if not sels:
            self.lbl_title.config(text="未选择")
            self.lbl_exe.config(text="")
            self.lbl_big.config(image="")
            self.lbl_big.image = None
            return
        exe = self.tree.item(sels[0], "values")[1]
        self.lbl_title.config(text=os.path.basename(exe))
        self.lbl_exe.config(text=exe)
        try:
            big = self._get_photo(exe, 48)
            self.lbl_big.config(image=big)
            self.lbl_big.image = big
        except Exception:
            self.lbl_big.config(image="")
            self.lbl_big.image = None

    def on_ok(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showinfo("提示", "请先选择至少一个进程")
            return
        exes = [self.tree.item(i, "values")[1] for i in sels]
        # 去重
        dedup = {}
        for e in exes:
            dedup[e] = True
        self.result = list(dedup.keys())
        self._close()

    def on_cancel(self):
        self.result = None
        self._close()

    def _close(self):
        self._closed = True
        if self._refresh_job:
            try:
                self.after_cancel(self._refresh_job)
            except Exception:
                pass
            self._refresh_job = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()
        # 释放图片缓存引用，避免内存泄漏
        self._photos.clear()
