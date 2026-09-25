import tkinter as tk
from tkinter import ttk, filedialog

from PIL import ImageTk, Image
from core.icons import get_exe_icon_pil
from core.paths import exe_key, app_name
from core.windows import taskbar_apps
from .theme import BG, INK, MUTED, ACCENT, FONT


class ProcessPicker(tk.Toplevel):
    """Launchpad-like picker showing running taskbar applications, not services."""

    def __init__(self, master, multiselect=True, existing=(), provider=taskbar_apps):
        super().__init__(master)
        self.title("添加游戏 · Galgame BGM")
        self.geometry("1080x780")
        self.minsize(840, 640)
        self.configure(bg=BG)
        self.result = None
        self.names = {}
        self.provider = provider
        self.multiselect = multiselect
        self.existing = {exe_key(e) for e in existing}
        self.selected = {}
        self._photos = {}
        self._items = []
        self._columns = 0
        self._filter_job = None
        self._refresh_job = None
        self._closed = False
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

        header = ttk.Frame(self, padding=(26, 20, 26, 12))
        header.pack(fill="x")
        ttk.Label(header, text="选择正在运行的游戏", font=(FONT, 20, "bold")).pack(anchor="w")
        ttk.Label(header, text="只显示任务栏应用 · 点击图标多选 · 已添加的游戏会自动监听",
                  style="Muted.TLabel").pack(anchor="w", pady=(5, 14))
        search = ttk.Frame(header)
        search.pack(fill="x")
        ttk.Label(search, text="搜索应用").pack(side="left", padx=(0, 10))
        self.var_q = tk.StringVar()
        self.entry = ttk.Entry(search, textvariable=self.var_q)
        self.entry.pack(side="left", fill="x", expand=True)
        self.var_q.trace_add("write", self._schedule_filter)
        ttk.Button(search, text="刷新", command=self.reload).pack(side="left", padx=(10, 0))

        body = ttk.Frame(self, padding=(20, 0, 20, 0))
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, background=BG, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        scroll.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.grid_frame = tk.Frame(self.canvas, bg=BG)
        self._window = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.canvas.bind("<Configure>", self._resize)
        self.grid_frame.bind("<Configure>", lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.bind("<MouseWheel>", self._wheel)
        self.bind("<Escape>", lambda _: self.on_cancel())
        self.bind("<Control-f>", lambda _: self.entry.focus_set())
        self.bind("<Return>", lambda _: self.on_ok())

        footer = ttk.Frame(self, padding=(26, 14, 26, 20))
        footer.pack(fill="x", side="bottom", before=body)
        self.summary = ttk.Label(footer, text="", style="Muted.TLabel")
        self.summary.pack(anchor="w", pady=(0, 10))
        ttk.Button(footer, text="找不到？从文件添加…", command=self.browse).pack(side="left")
        self.btn_ok = ttk.Button(footer, text="添加所选", style="Accent.TButton", command=self.on_ok)
        self.btn_ok.pack(side="right")
        ttk.Button(footer, text="取消", command=self.on_cancel).pack(side="right", padx=(0, 10))
        self.reload()
        self.transient(master)
        self.grab_set()
        self.entry.focus_set()

    def _wheel(self, event):
        self.canvas.yview_scroll(-int(event.delta / 120), "units")
        return "break"

    def _resize(self, event):
        self.canvas.itemconfigure(self._window, width=event.width)
        columns = max(3, event.width // 190)
        if columns != self._columns:
            self._columns = columns
            self.render()

    def _schedule_filter(self, *_):
        if self._filter_job:
            self.after_cancel(self._filter_job)
        self._filter_job = self.after(120, self.render)

    def reload(self):
        if self._closed:
            return
        if self._refresh_job:
            self.after_cancel(self._refresh_job)
        try:
            self._items = sorted(self.provider(), key=lambda a: (a.title or app_name(a.exe)).casefold())
            self._load_error = None
        except Exception:
            self._load_error = "暂时无法读取任务栏，请重试或从文件添加。"
        self.render()
        self._refresh_job = self.after(4000, self.reload)

    def render(self):
        self._filter_job = None
        focused = self.focus_get()
        focused_key = getattr(focused, "_app_key", None)
        for child in self.grid_frame.winfo_children():
            child.destroy()
        q = self.var_q.get().strip().casefold()
        items = [a for a in self._items if not q or q in (a.title + " " + app_name(a.exe)).casefold()]
        columns = self._columns or 5
        for index in range(8):
            self.grid_frame.columnconfigure(index, weight=1 if index < columns else 0, uniform="cards" if index < columns else "")
        for index, app in enumerate(items):
            key = exe_key(app.exe)
            known = key in self.existing
            selected = key in self.selected
            if key not in self._photos:
                self._photos[key] = ImageTk.PhotoImage(get_exe_icon_pil(app.exe).resize((56, 56), Image.Resampling.LANCZOS))
            title = app.title or app_name(app.exe)
            title = title if len(title) <= 14 else title[:13] + "…"
            suffix = "已添加" if known else ("✓ 已选择" if selected else app_name(app.exe))
            card = tk.Button(
                self.grid_frame, image=self._photos[key], text=f"{title}\n{suffix[:18]}",
                compound="top", wraplength=170, height=168, font=(FONT, 9), fg=MUTED if known else INK,
                bg="#e3eaff" if selected else "white", activebackground="#eaf0ff",
                relief="flat", bd=0, highlightthickness=2,
                highlightbackground=ACCENT if selected else BG, highlightcolor=ACCENT,
                padx=8, pady=12, cursor="hand2", takefocus=not known,
                state="disabled" if known else "normal", command=lambda a=app: self.toggle(a))
            card._app_key = key
            card.grid(row=index // columns, column=index % columns, sticky="nsew", padx=5, pady=5)
            if focused_key == key:
                card.focus_set()
        if not items:
            tk.Label(self.grid_frame, text=self._load_error or ("没有匹配的应用" if q else "未找到任务栏应用\n先打开游戏，或从文件添加"),
                     bg=BG, fg=MUTED, font=(FONT, 13), pady=65).grid(columnspan=columns, sticky="ew")
        self.summary.config(text=self._load_error or f"{len(self._items)} 个任务栏应用  ·  已选 {len(self.selected)} 个")
        self.btn_ok.config(text=f"添加所选（{len(self.selected)}）", state="normal" if self.selected else "disabled")

    def toggle(self, app):
        key = exe_key(app.exe)
        if key in self.selected:
            self.selected.pop(key)
        else:
            if not self.multiselect:
                self.selected.clear()
            self.selected[key] = app.exe
            self.names[key] = app.title or app_name(app.exe)
        self.render()

    def browse(self):
        files = filedialog.askopenfilenames(parent=self, title="选择游戏程序", filetypes=[("Windows 应用", "*.exe")])
        for path in files:
            key = exe_key(path)
            if key not in self.existing:
                self.selected[key] = path
                self.names[key] = app_name(path)
        self.render()

    def on_ok(self):
        if self.selected:
            self.result = list(self.selected.values())
            self._close()

    def on_cancel(self):
        self._close()

    def _close(self):
        self._closed = True
        for job in (self._refresh_job, self._filter_job):
            if job:
                self.after_cancel(job)
        self.grab_release()
        self.destroy()
