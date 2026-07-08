# =============================================================
#  ui/app.py — Ventana principal con pestanas
#  Modulo 1: Generador Archivo Plano PSL (Toledana del Sur)
#  Modulo 2: Conciliacion Proyectos (Fiduciaria Bancolombia)
# =============================================================

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .widgets import (
    ASSETS_DIR,
    COLOR_BG, COLOR_BLUE,
    FONT_TITLE,
)
from .tab_toledana    import TabToledana
from .tab_conciliacion import TabConciliacion


class App(tk.Tk):
    """Ventana principal unificada con ttk.Notebook."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Contabilidad - Toledana del Sur")
        self.configure(bg=COLOR_BG)
        self.resizable(True, True)
        self.minsize(780, 700)

        self._build_ui()
        self._load_favicon()

        self.update_idletasks()
        w  = max(self.winfo_width(),  820)
        h  = max(self.winfo_height(), 780)
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    # ── Construccion de la UI ─────────────────────────────────

    def _build_ui(self) -> None:
        # Header azul global
        hdr = tk.Frame(self, bg=COLOR_BLUE, pady=12)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="Contabilidad — Toledana del Sur",
            font=FONT_TITLE, bg=COLOR_BLUE, fg="white",
        ).pack()
        tk.Label(
            hdr, text="Archivo Plano PSL  |  Conciliacion Proyectos Fiduciaria",
            font=("Segoe UI", 9), bg=COLOR_BLUE, fg="#A8C4E0",
        ).pack()

        # Estilos del Notebook
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "TNotebook",
            background=COLOR_BG,
            borderwidth=0,
            tabmargins=[0, 4, 0, 0],
        )
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding=[18, 8],
            background="#D0D7E3",
            foreground=COLOR_BLUE,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#FFFFFF")],
            foreground=[("selected", COLOR_BLUE)],
            expand=[("selected", [1, 1, 1, 0])],
        )
        style.configure("TFrame", background=COLOR_BG)

        # Notebook
        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill="both", expand=True)

        # Pestanas
        self._tab_toledana     = TabToledana(self._notebook)
        self._tab_conciliacion = TabConciliacion(self._notebook)

        self._notebook.add(self._tab_toledana,     text="  Toledana del Sur  ")
        self._notebook.add(self._tab_conciliacion, text="  Conciliacion Proyectos  ")

        # Routing del scroll al canvas de la pestana activa
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event) -> None:
        """Enruta el scroll al canvas de la pestana activa."""
        tab_id = self._notebook.select()
        if not tab_id:
            return
        try:
            widget = self._notebook.nametowidget(tab_id)
            if hasattr(widget, "_canvas"):
                widget._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    # ── Favicon ───────────────────────────────────────────────

    def _load_favicon(self) -> None:
        ico = ASSETS_DIR / "favicon.ico"
        if ico.exists():
            try:
                self.iconbitmap(str(ico))
            except Exception:
                pass  # En Linux/Mac iconbitmap puede fallar; no es critico
