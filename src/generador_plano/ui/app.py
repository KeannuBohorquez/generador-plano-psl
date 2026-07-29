# =============================================================
#  ui/app.py — Ventana principal con selector de modulo
#  Modulo 1: Bancolombia Toledana  (Generador Archivo Plano PSL)
#  Modulo 2: Alianza               (Conciliacion Proyectos Fiduciaria)
# =============================================================

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .widgets import (
    ASSETS_DIR,
    COLOR_BG, COLOR_BLUE, COLOR_BORDER,
    FONT_TITLE, FONT_BOLD, FONT_LABEL,
)
from .tab_toledana     import TabToledana
from .tab_conciliacion import TabConciliacion

# Nombres visibles en el selector (en el orden en que aparecen)
NOMBRE_TOLEDANA     = "Bancolombia Toledana"
NOMBRE_CONCILIACION = "Alianza"


class App(tk.Tk):
    """Ventana principal unificada con selector desplegable de modulo."""

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

        # ── Selector de modulo ─────────────────────────────────
        sel = tk.Frame(self, bg="#E8EDF3", pady=10)
        sel.pack(fill="x")
        tk.Label(
            sel, text="Modulo:", font=FONT_BOLD,
            bg="#E8EDF3", fg=COLOR_BLUE,
        ).pack(side="left", padx=(16, 8))

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Modulo.TCombobox",
            fieldbackground="#FFFFFF",
            background="#FFFFFF",
            foreground=COLOR_BLUE,
            arrowcolor=COLOR_BLUE,
            bordercolor=COLOR_BORDER,
        )

        self._modulo_var = tk.StringVar(value=NOMBRE_TOLEDANA)
        self._combo = ttk.Combobox(
            sel, textvariable=self._modulo_var,
            values=[NOMBRE_TOLEDANA, NOMBRE_CONCILIACION],
            state="readonly", font=FONT_LABEL, width=28,
            style="Modulo.TCombobox",
        )
        self._combo.pack(side="left", ipady=3)
        self._combo.bind("<<ComboboxSelected>>", self._on_modulo_change)

        # ── Contenedor de modulos (apilados, se muestra uno a la vez) ──
        contenedor = tk.Frame(self, bg=COLOR_BG)
        contenedor.pack(fill="both", expand=True)
        contenedor.rowconfigure(0, weight=1)
        contenedor.columnconfigure(0, weight=1)

        self._tab_toledana     = TabToledana(contenedor)
        self._tab_conciliacion = TabConciliacion(contenedor)

        for frame in (self._tab_toledana, self._tab_conciliacion):
            frame.grid(row=0, column=0, sticky="nsew")

        self._frames = {
            NOMBRE_TOLEDANA:     self._tab_toledana,
            NOMBRE_CONCILIACION: self._tab_conciliacion,
        }
        self._activo = self._tab_toledana
        self._tab_toledana.tkraise()

        # Routing del scroll al canvas del modulo activo
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_modulo_change(self, event=None) -> None:
        nombre = self._modulo_var.get()
        frame = self._frames.get(nombre)
        if frame is None:
            return
        self._activo = frame
        frame.tkraise()

    def _on_mousewheel(self, event) -> None:
        """Enruta el scroll al canvas del modulo activo."""
        try:
            if hasattr(self._activo, "_canvas"):
                self._activo._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
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
