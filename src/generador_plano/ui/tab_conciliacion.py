# =============================================================
#  ui/tab_conciliacion.py — Pestana Conciliacion Proyectos
#  (Extracto Fiduciaria Bancolombia - BOSKE / 23LIVING / THECORNER)
# =============================================================

from __future__ import annotations

import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

from ..conciliacion.procesador import procesar, ResultadoConciliacion
from .widgets import (
    FilePicker, LogRedirect,
    COLOR_BG, COLOR_PANEL, COLOR_BLUE, COLOR_ACCENT, COLOR_GREEN, COLOR_BORDER,
    FONT_BOLD, FONT_LABEL, FONT_MONO, FONT_BTN,
)

MESES = [
    "01 - Enero",     "02 - Febrero",   "03 - Marzo",    "04 - Abril",
    "05 - Mayo",      "06 - Junio",     "07 - Julio",    "08 - Agosto",
    "09 - Septiembre","10 - Octubre",   "11 - Noviembre","12 - Diciembre",
]

# Nombre de mes a partir de la cadena del combobox
def _mes_nombre_de_seleccion(s: str) -> str:
    """Extrae 'Mayo' de '05 - Mayo' y lo devuelve en mayusculas."""
    return s.split(" - ")[1].upper() if " - " in s else s.upper()


class TabConciliacion(tk.Frame):
    """
    Pestana para procesar el extracto PDF de la fiduciaria
    Bancolombia y generar el archivo plano de conciliacion
    de proyectos.
    """

    def __init__(self, parent, **kw) -> None:
        super().__init__(parent, bg=COLOR_BG, **kw)
        self._build_ui()
        self._set_defaults()

    # ── Construccion de la UI ─────────────────────────────────

    def _section(self, parent, title: str) -> tk.LabelFrame:
        f = tk.LabelFrame(
            parent, text=f"  {title}  ", font=FONT_BOLD,
            bg=COLOR_PANEL, fg=COLOR_BLUE, relief="flat",
            highlightbackground=COLOR_BORDER, highlightthickness=1,
            padx=14, pady=10,
        )
        f.pack(fill="x", padx=16, pady=(0, 8))
        return f

    def _build_ui(self) -> None:
        # Canvas con scroll
        self._canvas = tk.Canvas(self, bg=COLOR_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(self._canvas, bg=COLOR_BG)
        win  = self._canvas.create_window((0, 0), window=body, anchor="nw")
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfig(win, width=e.width),
        )
        body.bind(
            "<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )

        self._build_form(body)

    def _build_form(self, body: tk.Frame) -> None:
        # ── Periodo ───────────────────────────────────────────
        sp = self._section(body, "Periodo")
        rp = tk.Frame(sp, bg=COLOR_PANEL)
        rp.pack(fill="x")
        tk.Label(rp, text="Mes:", font=FONT_BOLD, bg=COLOR_PANEL, fg=COLOR_BLUE).pack(side="left")
        self.mes_var = tk.StringVar()
        ttk.Combobox(
            rp, textvariable=self.mes_var, values=MESES,
            state="readonly", width=22, font=FONT_LABEL,
        ).pack(side="left", padx=(6, 24), ipady=3)
        tk.Label(rp, text="Ano:", font=FONT_BOLD, bg=COLOR_PANEL, fg=COLOR_BLUE).pack(side="left")
        self.anio_var = tk.StringVar(value="2026")
        tk.Spinbox(
            rp, from_=2020, to=2040, textvariable=self.anio_var,
            width=7, font=FONT_LABEL, relief="flat",
            highlightthickness=1, highlightbackground=COLOR_BORDER,
        ).pack(side="left", padx=(6, 0), ipady=3)

        # ── Archivos de entrada ───────────────────────────────
        sf = self._section(body, "Archivos de entrada")
        self.pdf_picker = FilePicker(
            sf, "1.  Extracto fiduciaria  (PDF — sin contrasena)",
            [("PDF", "*.pdf"), ("Todos", "*.*")])
        self.pdf_picker.pack(fill="x", pady=(0, 10))
        self.clientes_picker = FilePicker(
            sf, "2.  Tabla de clientes  (Excel — hoja CONSOLIDADO ALIANZA)",
            [("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        self.clientes_picker.pack(fill="x")

        # ── Archivo de salida ─────────────────────────────────
        so = self._section(body, "Archivo de salida")
        self.out_picker = FilePicker(
            so, "Guardar resultado como...",
            [("Excel 97-2003", "*.xls")], save=True, defaultextension=".xls")
        self.out_picker.pack(fill="x")

        # ── Boton generar ─────────────────────────────────────
        bf = tk.Frame(body, bg=COLOR_BG)
        bf.pack(fill="x", padx=16, pady=(4, 8))
        self.btn_gen = tk.Button(
            bf, text="  Procesar Extracto Fiduciario  ", font=FONT_BTN,
            bg=COLOR_GREEN, fg="white", activebackground="#155724",
            activeforeground="white", relief="flat", cursor="hand2",
            padx=20, pady=11, command=self._on_generar)
        self.btn_gen.pack(fill="x")

        # ── Barra de estado ───────────────────────────────────
        self.status_var = tk.StringVar(value="Listo")
        self.status_bar = tk.Label(
            body, textvariable=self.status_var, font=("Segoe UI", 9),
            bg="#E8EDF3", fg=COLOR_BLUE, anchor="w", padx=10, pady=4)
        self.status_bar.pack(fill="x", padx=16, pady=(0, 4))

        # ── Log terminal ──────────────────────────────────────
        log_frame = tk.LabelFrame(
            body, text="  Registro de ejecucion  ", font=FONT_BOLD,
            bg=COLOR_BG, fg=COLOR_BLUE, relief="flat",
            highlightbackground=COLOR_BORDER, highlightthickness=1,
            padx=12, pady=8)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        ltb = tk.Frame(log_frame, bg=COLOR_BG)
        ltb.pack(fill="x", pady=(0, 4))
        tk.Button(ltb, text="Copiar log", font=("Segoe UI", 8), relief="flat",
                  bg="#DEE3EB", cursor="hand2", padx=8, pady=2,
                  command=self._copy_log).pack(side="right", padx=(4, 0))
        tk.Button(ltb, text="Limpiar", font=("Segoe UI", 8), relief="flat",
                  bg="#DEE3EB", cursor="hand2", padx=8, pady=2,
                  command=self._log_clear).pack(side="right")

        self.log = tk.Text(
            log_frame, height=14, font=FONT_MONO,
            bg="#0D1117", fg="#C9D1D9", relief="flat",
            state="disabled", wrap="word", padx=10, pady=8)
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

        self.log.tag_config("ok",   foreground="#3FB950")
        self.log.tag_config("err",  foreground="#F85149")
        self.log.tag_config("warn", foreground="#E3B341")
        self.log.tag_config("info", foreground="#79C0FF")
        self.log.tag_config("head", foreground="#E3B341", font=("Consolas", 9, "bold"))
        self.log.tag_config("dim",  foreground="#8B949E")
        self.log.tag_config("val",  foreground="#D2A8FF")

    # ── Helpers de UI ─────────────────────────────────────────

    def _set_defaults(self) -> None:
        import datetime
        hoy = datetime.date.today()
        self.mes_var.set(MESES[max(hoy.month - 2, 0)])
        self.anio_var.set(str(hoy.year))

    def _log(self, msg: str, tag: str = "") -> None:
        def _do():
            self.log.config(state="normal")
            self.log.insert("end", msg + "\n", tag)
            self.log.see("end")
            self.log.config(state="disabled")
        self.after(0, _do)

    def _log_clear(self) -> None:
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _copy_log(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.log.get("1.0", "end"))
        self.status_var.set("Log copiado al portapapeles")
        self.after(2000, lambda: self.status_var.set("Listo"))

    def _set_status(self, msg: str, color: str | None = None) -> None:
        self.after(0, lambda: self.status_var.set(msg))
        self.after(0, lambda: self.status_bar.config(fg=color or COLOR_BLUE))

    def _mes_seleccionado(self) -> tuple[str | None, str | None]:
        """Retorna (mes_nombre_upper, anio_str) o (None, None) si no hay seleccion."""
        s = self.mes_var.get()
        if not s:
            return None, None
        return _mes_nombre_de_seleccion(s), self.anio_var.get()

    # ── Logica de procesamiento ───────────────────────────────

    def _on_generar(self) -> None:
        mes, anio = self._mes_seleccionado()
        errores = []
        if not mes:                           errores.append("Selecciona un mes y ano.")
        if not self.pdf_picker.get():         errores.append("Falta el extracto fiduciario (PDF).")
        if not self.clientes_picker.get():    errores.append("Falta la tabla de clientes (Excel).")
        if not self.out_picker.get():         errores.append("Indica donde guardar el archivo de salida.")
        if errores:
            messagebox.showerror("Campos incompletos",
                                 "\n".join(f"  * {e}" for e in errores))
            return
        self.btn_gen.config(state="disabled", text="Procesando...")
        self._set_status("Procesando...", COLOR_ACCENT)
        self._log_clear()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self) -> None:
        mes_nombre, anio = self._mes_seleccionado()
        ruta_pdf      = self.pdf_picker.get()
        ruta_clientes = self.clientes_picker.get()
        ruta_salida   = self.out_picker.get()

        L = self._log

        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = LogRedirect(L, "dim")
        sys.stderr = LogRedirect(L, "err")

        try:
            L("=" * 54, "dim")
            L(f"  CONCILIACION PROYECTOS  |  {mes_nombre.capitalize()} {anio}", "head")
            L("=" * 54, "dim")

            # 1. Leer PDF
            L(f"\n[1/3]  Leyendo extracto fiduciario...", "info")
            L(f"       Archivo: {Path(ruta_pdf).name}", "dim")

            # 2. Leer clientes
            L(f"\n[2/3]  Cruzando con tabla de clientes...", "info")
            L(f"       Archivo: {Path(ruta_clientes).name}", "dim")

            # 3. Procesar
            L(f"\n[3/3]  Calculando archivo plano...", "info")

            try:
                resultado: ResultadoConciliacion = procesar(
                    ruta_pdf=ruta_pdf,
                    ruta_clientes=ruta_clientes,
                    mes_nombre=mes_nombre,
                    anio=anio,
                    ruta_salida=ruta_salida,
                )
            except PermissionError:
                L("\n  ERROR: No se pudo guardar el archivo.", "err")
                L("  Cierra el archivo si esta abierto en Excel.", "err")
                raise
            except Exception as ex:
                L(f"\n  ERROR: {ex}", "err")
                raise

            # Resumen
            r = resultado
            L("\n" + "=" * 54, "dim")
            L("  ARCHIVO GENERADO EXITOSAMENTE", "ok")
            L("=" * 54, "dim")
            L(f"  Ruta:     {ruta_salida}", "dim")
            L(f"  Filas:    {r.n_filas}", "dim")
            L(f"  Proyecto: {r.proyecto}", "val")
            L("─" * 54, "dim")
            L("  PARAMETROS DEL PLANO", "head")
            L(f"  {'Compania':<20}  {r.compania}", "")
            L(f"  {'Division':<20}  {r.division}", "")
            L(f"  {'Centro':<20}  {r.centro}", "")
            L(f"  {'Periodo':<20}  {mes_nombre.capitalize()} {anio}", "")
            L("─" * 54, "dim")

            if r.proyecto == "DESCONOCIDO":
                L("\n  AVISO: No se detecto el proyecto (BOSKE/23LIVING/THECORNER).", "warn")
                L("  Verifica que el PDF corresponde al extracto fiduciario.", "warn")
                L("  Compania/Division/Centro quedaron en 00/00/000.", "warn")

            L("\n  RECUERDA asignar el N de comprobante en PSL", "warn")

            self._set_status(
                f"Listo  |  {r.n_filas} filas  |  Proyecto: {r.proyecto}  |  {Path(ruta_salida).name}",
                COLOR_GREEN,
            )

            self.after(0, lambda:
                messagebox.showinfo(
                    "Exito",
                    f"Conciliacion generada correctamente.\n\n"
                    f"Proyecto: {r.proyecto}\n"
                    f"Filas: {r.n_filas}\n\n"
                    f"Archivo: {ruta_salida}",
                ))

        except Exception as ex:
            msg = str(ex)
            tb  = traceback.format_exc()
            L("\n" + "!" * 54, "err")
            L("  ERROR DURANTE EL PROCESAMIENTO", "err")
            L("!" * 54, "err")
            L(tb, "err")
            self._set_status(f"Error: {msg}", "#C0392B")
            self.after(0, lambda m=msg: messagebox.showerror(
                "Error",
                f"Ocurrio un error:\n\n{m}\n\nRevisa el Registro para mas detalles.",
            ))
        finally:
            sys.stdout, sys.stderr = old_out, old_err
            self.after(0, lambda: self.btn_gen.config(
                state="normal", text="  Procesar Extracto Fiduciario  "))
