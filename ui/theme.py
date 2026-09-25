import tkinter as tk
from tkinter import ttk

BG = "#f4f6fc"
INK = "#17233e"
MUTED = "#64718b"
ACCENT = "#4568df"
FONT = "Microsoft YaHei UI"
MODE_LABELS = {"not_foreground": "切到后台时静音", "minimized_only": "仅最小化时静音"}


def configure(root):
    root.configure(bg=BG)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", font=(FONT, 10))
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=INK)
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("Title.TLabel", font=(FONT, 23, "bold"))
    style.configure("TButton", padding=(14, 8))
    style.configure("Accent.TButton", background=ACCENT, foreground="white", borderwidth=0)
    style.map("Accent.TButton", background=[("active", "#3555c5"), ("disabled", "#a7b1d1")])
    style.configure("Treeview", background="white", fieldbackground="white", foreground=INK,
                    rowheight=60, borderwidth=0, font=(FONT, 10))
    style.configure("Treeview.Heading", background=BG, foreground=MUTED, padding=9)
    style.map("Treeview", background=[("selected", "#e3eaff")], foreground=[("selected", INK)])


def label(parent, text, **kwargs):
    return tk.Label(parent, text=text, bg=BG, fg=INK, **kwargs)
