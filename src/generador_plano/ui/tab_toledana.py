# =============================================================
#  ui/tab_toledana.py — Pestana Generador Archivo Plano PSL
#  (P.A. Toledana del Sur - Conciliacion BanColombia)
# =============================================================

from __future__ import annotations

import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

from ..config import PASSWORD_DEFAULT, NOMBRES_MES_LARGO, NOMBRES_MES_CORTO
from ..extractor import parsear_extracto
from ..lector import leer_propietarios, leer_sin_conciliar
from ..constructor import construir_filas
from ..exportador import exportar_excel
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


def _nombre_mes(mes: int) -> str:
    return NOMBRES_MES_LARGO[mes]

def _num_doc(anio: int, mes: int) -> str:
    return f"{NOMBRES_MES_CORTO[mes]}{anio}"


class TabToledana(tk.Frame):
    """
    Pestana del Generador de Archivo Plano PSL para
    P.A. Toledana del Sur.
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
            sf, "1.  Extracto bancario  (PDF con contrasena)",
            [("PDF", "*.pdf"), ("Todos", "*.*")])
        self.pdf_picker.pack(fill="x", pady=(0, 10))
        self.mov_picker = FilePicker(
            sf, "2.  Informe de movimientos por propietario  (Excel protegido)",
            [("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        self.mov_picker.pack(fill="x", pady=(0, 10))
        self.sin_conc_picker = FilePicker(
            sf, "3.  Movimientos sin conciliar  (Excel)",
            [("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        self.sin_conc_picker.pack(fill="x")

        # ── Archivo de salida ─────────────────────────────────
        so = self._section(body, "Archivo de salida")
        self.out_picker = FilePicker(
            so, "Guardar archivo plano como...",
            [("Excel", "*.xlsx")], save=True)
        self.out_picker.pack(fill="x")

        # ── Opciones ──────────────────────────────────────────
        sop = self._section(body, "Opciones")
        grid = tk.Frame(sop, bg=COLOR_PANEL)
        grid.pack(fill="x")

        tk.Label(grid, text="Contrasena PDF:", font=FONT_BOLD,
                 bg=COLOR_PANEL, fg=COLOR_BLUE).grid(row=0, column=0, sticky="w", pady=4)
        self.pwd_pdf = tk.Entry(
            grid, show="*", font=FONT_LABEL, width=20, relief="flat",
            highlightthickness=1, highlightbackground=COLOR_BORDER, bg="#F8FAFC")
        self.pwd_pdf.grid(row=0, column=1, sticky="w", padx=(8, 24), ipady=4)

        tk.Label(grid, text="Contrasena Excel:", font=FONT_BOLD,
                 bg=COLOR_PANEL, fg=COLOR_BLUE).grid(row=0, column=2, sticky="w")
        self.pwd_xls = tk.Entry(
            grid, show="*", font=FONT_LABEL, width=20, relief="flat",
            highlightthickness=1, highlightbackground=COLOR_BORDER, bg="#F8FAFC")
        self.pwd_xls.grid(row=0, column=3, sticky="w", padx=(8, 0), ipady=4)

        self.show_pwd = tk.BooleanVar()
        tk.Checkbutton(
            grid, text="Mostrar contrasenas", variable=self.show_pwd,
            bg=COLOR_PANEL, font=FONT_LABEL, fg="#555",
            activebackground=COLOR_PANEL, command=self._toggle_pwd,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        tk.Label(grid, text="N comprobante:", font=FONT_BOLD,
                 bg=COLOR_PANEL, fg=COLOR_BLUE).grid(row=1, column=2, sticky="w", pady=(4, 0))
        self.n_comp = tk.Entry(
            grid, font=FONT_LABEL, width=20, relief="flat",
            highlightthickness=1, highlightbackground=COLOR_BORDER, bg="#F8FAFC")
        self.n_comp.grid(row=1, column=3, sticky="w", padx=(8, 0), ipady=4, pady=(4, 0))
        tk.Label(grid, text="(vacio = asignar en PSL)", font=("Segoe UI", 8),
                 fg="#888", bg=COLOR_PANEL).grid(row=2, column=3, sticky="w", padx=(8, 0))

        # ── Boton generar ─────────────────────────────────────
        bf = tk.Frame(body, bg=COLOR_BG)
        bf.pack(fill="x", padx=16, pady=(4, 8))
        self.btn_gen = tk.Button(
            bf, text="  Generar Archivo Plano  ", font=FONT_BTN,
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

        self.log.tag_config("ok",       foreground="#3FB950")
        self.log.tag_config("err",      foreground="#F85149")
        self.log.tag_config("warn",     foreground="#E3B341")
        self.log.tag_config("info",     foreground="#79C0FF")
        self.log.tag_config("head",     foreground="#E3B341", font=("Consolas", 9, "bold"))
        self.log.tag_config("dim",      foreground="#8B949E")
        self.log.tag_config("val",      foreground="#D2A8FF")
        self.log.tag_config("faltante", foreground="#FF0000",
                            font=("Consolas", 9, "bold"), background="#2D0000")

    # ── Helpers de UI ─────────────────────────────────────────

    def _toggle_pwd(self) -> None:
        ch = "" if self.show_pwd.get() else "*"
        self.pwd_pdf.config(show=ch)
        self.pwd_xls.config(show=ch)

    def _set_defaults(self) -> None:
        import datetime
        hoy = datetime.date.today()
        self.mes_var.set(MESES[max(hoy.month - 2, 0)])
        self.anio_var.set(str(hoy.year))
        self.pwd_pdf.insert(0, PASSWORD_DEFAULT)
        self.pwd_xls.insert(0, PASSWORD_DEFAULT)

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

    def _mes_anio(self) -> tuple[int | None, int | None]:
        s = self.mes_var.get()
        if not s:
            return None, None
        return int(self.anio_var.get()), int(s.split(" - ")[0])

    # ── Logica de generacion ──────────────────────────────────

    def _on_generar(self) -> None:
        anio, mes = self._mes_anio()
        errores = []
        if not anio:                        errores.append("Selecciona un mes y ano.")
        if not self.pdf_picker.get():       errores.append("Falta el extracto bancario (PDF).")
        if not self.mov_picker.get():       errores.append("Falta el informe de movimientos.")
        if not self.sin_conc_picker.get():  errores.append("Falta el archivo sin conciliar.")
        if not self.out_picker.get():       errores.append("Indica donde guardar el archivo de salida.")
        if errores:
            messagebox.showerror("Campos incompletos",
                                 "\n".join(f"  * {e}" for e in errores))
            return
        self.btn_gen.config(state="disabled", text="Procesando...")
        self._set_status("Procesando...", COLOR_ACCENT)
        self._log_clear()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self) -> None:
        anio, mes = self._mes_anio()
        extracto  = self.pdf_picker.get()
        informe   = self.mov_picker.get()
        sin_conc  = self.sin_conc_picker.get()
        salida    = self.out_picker.get()
        pwd_pdf   = self.pwd_pdf.get()
        pwd_xls   = self.pwd_xls.get()
        n_comp    = self.n_comp.get().strip()

        L = self._log

        # Redirigir prints sueltos al log mientras corre
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = LogRedirect(L, "dim")
        sys.stderr = LogRedirect(L, "err")

        try:
            L("=" * 54, "dim")
            L(f"  GENERADOR ARCHIVO PLANO  |  {_nombre_mes(mes)} {anio}", "head")
            L("=" * 54, "dim")

            # 1. Extracto PDF
            L(f"\n[1/4]  Parseando extracto bancario...", "info")
            L(f"       Archivo: {Path(extracto).name}", "dim")
            try:
                df_ext = parsear_extracto(extracto, pwd_pdf)
            except Exception as ex:
                L(f"\n  ERROR al leer el PDF:", "err")
                if "password" in str(ex).lower() or "incorrect" in str(ex).lower():
                    L("  La contrasena del PDF es incorrecta.", "err")
                else:
                    L(f"  {ex}", "err")
                raise

            abonos = df_ext[df_ext["valor"] > 0]["valor"].sum()
            cargos = abs(df_ext[df_ext["valor"] < 0]["valor"].sum())
            L(f"       {len(df_ext)} transacciones encontradas", "ok")
            L(f"       Abonos:  ${abonos:>20,.0f}", "val")
            L(f"       Cargos:  ${cargos:>20,.0f}", "val")
            L("       Desglose por tipo:", "dim")
            for tipo, cnt in df_ext["tipo"].value_counts().items():
                sub = df_ext[df_ext["tipo"] == tipo]["valor"]
                L(f"         {tipo:<18} {cnt:>3} mov  |  ${sub.sum():>15,.0f}", "dim")

            sin_cls = df_ext[df_ext["tipo"] == "OTRO"]
            if len(sin_cls) > 0:
                L(f"\n  AVISO: {len(sin_cls)} transacciones sin clasificar (OTRO):", "warn")
                for _, r in sin_cls.iterrows():
                    L(f"    {r['fecha']}  {r['desc'][:50]}  ${r['valor']:,.0f}", "warn")

            # 2. Informe movimientos
            L(f"\n[2/4]  Leyendo informe de movimientos...", "info")
            L(f"       Archivo: {Path(informe).name}", "dim")
            try:
                df_prop = leer_propietarios(informe, pwd_xls, anio, mes)
            except Exception as ex:
                L(f"\n  ERROR al leer el informe:", "err")
                if "password" in str(ex).lower() or "incorrect" in str(ex).lower():
                    L("  La contrasena del Excel es incorrecta.", "err")
                elif "Mov_Por_Propietario" in str(ex):
                    L("  No se encontro la hoja 'Mov_Por_Propietario'.", "err")
                else:
                    L(f"  {ex}", "err")
                raise

            total_prop = df_prop["Valor"].sum()
            L(f"       {len(df_prop)} propietarios identificados", "ok")
            L(f"       Total aportes:  ${total_prop:>17,.0f}", "val")

            # 3. Sin conciliar
            L(f"\n[3/4]  Leyendo movimientos sin conciliar...", "info")
            L(f"       Archivo: {Path(sin_conc).name}", "dim")
            try:
                total_sin_conc, n_sin_conc = leer_sin_conciliar(
                    sin_conc, anio, mes, pwd=pwd_xls)
            except Exception as ex:
                L(f"\n  ERROR al leer sin conciliar:", "err")
                L(f"  {ex}", "err")
                raise

            recaudo_ext = df_ext[df_ext["tipo"] == "RECAUDO"]["valor"].sum()
            diferencia  = recaudo_ext - total_prop
            L(f"       {n_sin_conc} registros sin conciliar para {_nombre_mes(mes)} {anio}", "ok")
            L(f"       Total sin conciliar (archivo):  ${total_sin_conc:>12,.0f}", "val")
            L(f"       Diferencia calculada (ext-ap):  ${diferencia:>12,.0f}", "val")

            if n_sin_conc == 0 and diferencia > 1:
                L(f"\n  AVISO: No hay registros sin conciliar para este mes.", "warn")
                L(f"  Diferencia calculada: ${diferencia:,.0f}", "warn")
                L("  Verifica que el archivo corresponde al periodo.", "warn")
            elif abs(total_sin_conc - diferencia) > 1:
                L(f"\n  AVISO: Archivo sin conciliar (${total_sin_conc:,.0f})", "warn")
                L(f"  no coincide con diferencia calculada (${diferencia:,.0f}).", "warn")
                L(f"  Diferencia entre ambos: ${abs(total_sin_conc - diferencia):,.0f}", "warn")
                L("  Se usara el valor del archivo sin conciliar.", "warn")
            else:
                L("       Conciliacion OK: valores coinciden", "ok")

            # 4. Construir y exportar
            L(f"\n[4/4]  Generando archivo plano...", "info")
            sc = total_sin_conc if total_sin_conc > 0 else None
            resultado = construir_filas(df_ext, df_prop, anio, mes, n_comp, sc)

            for adv in resultado.advertencias:
                L(f"\n  AVISO: {adv}", "warn")

            # Alerta de dinero faltante
            if resultado.dinero_faltante > 1:
                L("\n" + "!" * 54, "faltante")
                L("  *** ALERTA: DINERO FALTANTE ***", "faltante")
                L("!" * 54, "faltante")
                L(f"  Monto faltante:  ${resultado.dinero_faltante:>18,.2f}", "faltante")
                L(f"  Recaudo extracto:  ${resultado.recaudo:>16,.2f}", "faltante")
                L(f"  Aportes identif.:  ${resultado.propietarios_total:>16,.2f}", "faltante")
                L(f"  Sin conciliar:     ${resultado.recaudos_pend:>16,.2f}", "faltante")
                L(f"  Diferencia:        ${resultado.dinero_faltante:>16,.2f}", "faltante")
                L("!" * 54, "faltante")
                L("  La fila DINERO FALTANTE fue agregada al archivo plano", "faltante")
                L("  en ROJO para que quede visible en PSL.", "faltante")

            try:
                exportar_excel(resultado.filas, salida,
                               indices_faltante=resultado.indices_faltante)
            except PermissionError:
                L("\n  ERROR: No se pudo guardar el archivo.", "err")
                L("  Cierra el archivo si esta abierto en Excel.", "err")
                raise

            # Resumen final
            r = resultado
            L("\n" + "=" * 54, "dim")
            L("  ARCHIVO GENERADO EXITOSAMENTE", "ok")
            L("=" * 54, "dim")
            L(f"  Ruta:   {salida}", "dim")
            L(f"  Filas:  {len(r.filas)}", "dim")
            L("─" * 54, "dim")
            L("  RESUMEN FINANCIERO", "head")
            L(f"  {'Recaudo extracto':<34}  ${r.recaudo:>14,.2f}", "")
            L(f"  {'Gastos bancarios':<34}  ${r.gastos:>14,.2f}", "")
            L(f"  {'Aportes propietarios':<34}  ${r.propietarios_total:>14,.2f}", "")
            if r.recaudos_pend > 0:
                L(f"  {'Sin conciliar (archivo)':<34}  ${r.recaudos_pend:>14,.2f}", "warn")
                if abs(r.recaudos_pend - r.diferencia_calc) > 1:
                    L(f"  {'  dif. calculada':<34}  ${r.diferencia_calc:>14,.2f}", "dim")
            if r.dinero_faltante > 1:
                L(f"  {'*** DINERO FALTANTE ***':<34}  ${r.dinero_faltante:>14,.2f}", "faltante")
            L(f"  {'DB banco (neto)':<34}  ${r.db_banco:>14,.2f}", "")
            L(f"  {'Traslado fiducia':<34}  ${r.traslado_neto:>14,.2f}", "")
            L(f"  {'N propietarios':<34}  {r.n_propietarios:>15}", "")
            L("─" * 54, "dim")
            if not n_comp:
                L("\n  RECUERDA asignar el N de comprobante en PSL", "warn")

            hay_faltante = r.dinero_faltante > 1
            status_color = "#C0392B" if hay_faltante else COLOR_GREEN
            status_msg   = (
                f"ALERTA: DINERO FALTANTE ${r.dinero_faltante:,.0f}  |  {Path(salida).name}"
                if hay_faltante else
                f"Listo  |  {len(r.filas)} filas  |  {r.n_propietarios} propietarios  |  {Path(salida).name}"
            )
            self._set_status(status_msg, status_color)

            n_filas   = len(r.filas)
            n_prop    = r.n_propietarios
            faltante  = r.dinero_faltante

            if hay_faltante:
                self.after(0, lambda s=salida, n=n_filas, p=n_prop, f=faltante:
                    messagebox.showwarning(
                        "ALERTA - Dinero Faltante",
                        f"El archivo fue generado pero hay DINERO FALTANTE.\n\n"
                        f"Monto faltante:  ${f:,.2f}\n\n"
                        f"Se agrego una fila en ROJO al archivo plano con el comentario:\n"
                        f"'*** DINERO FALTANTE POR IDENTIFICAR ***'\n\n"
                        f"Revisa el extracto bancario para identificar los movimientos faltantes.\n\n"
                        f"Archivo: {s}\nFilas: {n}  |  Propietarios: {p}",
                    ))
            else:
                self.after(0, lambda s=salida, n=n_filas, p=n_prop:
                    messagebox.showinfo(
                        "Exito",
                        f"Archivo generado correctamente.\n\n"
                        f"Filas: {n}\nPropietarios: {p}\n\n{s}",
                    ))

        except Exception as ex:
            msg = str(ex)
            tb  = traceback.format_exc()
            L("\n" + "!" * 54, "err")
            L("  ERROR DURANTE LA GENERACION", "err")
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
                state="normal", text="  Generar Archivo Plano  "))
