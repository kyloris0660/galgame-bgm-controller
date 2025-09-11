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

        # 顶栏：搜索 + “优先显示推荐”
        bar = ttk.Frame(self, padding=(10, 8, 10, 6))
        bar.pack(fill="x")
        ttk.Label(bar, text="搜索：", font=("", 10, "bold")).pack(side="left")
        self.var_q = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.var_q, width=40)
        ent.pack(side="left", fill="x", expand=True, padx=(4, 8))
        ent.bind("<KeyRelease>", self._refresh)
        self.var_hot = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            bar, text="优先显示推荐", variable=self.var_hot, command=self._refresh
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

        self._load()
        self._refresh()
        self.transient(master)
        self.grab_set()
        ent.focus_set()

    def _get_foreground_exe(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
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
        for p in psutil.process_iter(["name", "exe"]):
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
        # 预热图标
        for it in sorted(self._items, key=lambda x: -x["score"])[:12]:
            self._get_photo(it["exe"], 20)

    def _get_photo(self, exe, size):
        if (exe, size) in self._photos:
            return self._photos[(exe, size)]
        pil = get_exe_icon_pil(exe, large=True)
        if pil.size != (size, size):
            pil = pil.resize((size, size), resample=1)
        ph = ImageTk.PhotoImage(pil)
        self._photos[(exe, size)] = ph
        return ph

    def _refresh(self, *_):
        q = (self.var_q.get() or "").lower()
        only_hot = self.var_hot.get()
        items = [
            it
            for it in self._items
            if (not q or q in it["name"].lower() or q in it["exe"].lower())
        ]
        items.sort(key=lambda x: (-x["score"], x["name"].lower()))
        self.tree.delete(*self.tree.get_children())
        for it in items:
            tags = ("hot",) if it["score"] >= 100 else ()
            if only_hot and not tags:
                # 调暗显示而非隐藏
                iid = self.tree.insert(
                    "",
                    "end",
                    text="",
                    image=self._get_photo(it["exe"], 18),
                    values=(it["name"], it["exe"]),
                )
                self.tree.item(iid, tags=("dim",))
                self.tree.tag_configure("dim", foreground="#666666")
            else:
                self.tree.insert(
                    "",
                    "end",
                    text="",
                    image=self._get_photo(it["exe"], 18),
                    values=(it["name"], it["exe"]),
                    tags=tags,
                )
        self._refresh_ok()

    def _refresh_ok(self, *_):
        has_sel = bool(self.tree.selection())
        self.btn_ok.configure(state=("normal" if has_sel else "disabled"))

    def _on_select(self, *_):
        self._refresh_ok()
        sels = self.tree.selection()
        if not sels:
            self.lbl_title.config(text="未选择")
            self.lbl_exe.config(text="")
            self.lbl_big.config(image="")
            return
        exe = self.tree.item(sels[0], "values")[1]
        self.lbl_title.config(text=os.path.basename(exe))
        self.lbl_exe.config(text=exe)
        big = self._get_photo(exe, 48)
        self.lbl_big.config(image=big)
        self.lbl_big.image = big

    def on_ok(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showinfo("提示", "请先选择至少一个进程")
            return
        exes = [self.tree.item(i, "values")[1] for i in sels]
        self.result = exes
        self.grab_release()
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.grab_release()
        self.destroy()
