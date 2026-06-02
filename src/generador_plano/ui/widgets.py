# =============================================================
#  ui/widgets.py — Componentes reutilizables de la UI
# =============================================================

import sys
import tkinter as tk
from tkinter import filedialog
from pathlib import Path


# ── Colores y fuentes ─────────────────────────────────────────
COLOR_BG     = "#F0F4F8"
COLOR_PANEL  = "#FFFFFF"
COLOR_BLUE   = "#1F4E79"
COLOR_ACCENT = "#2E75B6"
COLOR_GREEN  = "#1E7E34"
COLOR_BORDER = "#D0D7E3"

FONT_LABEL = ("Segoe UI", 9)
FONT_BOLD  = ("Segoe UI", 9, "bold")
FONT_MONO  = ("Consolas", 9)
FONT_BTN   = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 13, "bold")

# Carpeta base: junto al .exe si esta compilado, o junto a main.py
BASE_DIR = (
    Path(sys.executable).parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parents[3]
)


class LogRedirect:
    """Redirige stdout/stderr al widget de log de la UI."""

    def __init__(self, callback, tag: str = "dim"):
        self.callback = callback
        self.tag = tag

    def write(self, msg: str) -> None:
        if msg.strip():
            self.callback(msg.rstrip(), self.tag)

    def flush(self) -> None:
        pass


class FilePicker(tk.Frame):
    """Widget de seleccion de archivo con etiqueta y boton Examinar."""

    def __init__(self, parent, label: str, filetypes: list, save: bool = False, **kw):
        super().__init__(parent, bg=COLOR_PANEL, **kw)
        self.save = save
        self.filetypes = filetypes

        tk.Label(
            self, text=label, font=FONT_BOLD, bg=COLOR_PANEL,
            fg=COLOR_BLUE, anchor="w",
        ).pack(fill="x", pady=(0, 2))

        row = tk.Frame(self, bg=COLOR_PANEL)
        row.pack(fill="x")

        self.var = tk.StringVar()
        tk.Entry(
            row, textvariable=self.var, font=FONT_LABEL, relief="flat",
            highlightthickness=1, highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_ACCENT, bg="#F8FAFC",
        ).pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))

        tk.Button(
            row, text="Examinar", font=FONT_LABEL, bg=COLOR_ACCENT, fg="white",
            relief="flat", activebackground=COLOR_BLUE, activeforeground="white",
            cursor="hand2", padx=10, pady=4, command=self._browse,
        ).pack(side="right")

    def _browse(self) -> None:
        kw = dict(
            defaultextension=".xlsx",
            filetypes=self.filetypes,
            initialdir=str(BASE_DIR),
        )
        path = (
            filedialog.asksaveasfilename(**kw)
            if self.save
            else filedialog.askopenfilename(**kw)
        )
        if path:
            self.var.set(path)

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, value: str) -> None:
        self.var.set(value)
