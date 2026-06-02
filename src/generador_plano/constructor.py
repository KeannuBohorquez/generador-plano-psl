# =============================================================
#  constructor.py — Construccion de las filas del archivo plano
# =============================================================

from __future__ import annotations

from calendar import monthrange
from dataclasses import dataclass, field
from datetime import date

import pandas as pd

from .config import (
    COMPANIA, DIVISION, CENTRO, FUENTE, MONEDA, ORIGEN,
    NIT_BANCOLOMBIA,
    CTA_COMISIONES, CTA_GMF, CTA_IVA_COMISION,
    CTA_PROPIETARIOS, CTA_PEND_IDENTIF,
    CTA_BANCO_CTE, CTA_FIDUCIA,
    NOMBRES_MES_CORTO, NOMBRES_MES_LARGO,
)


# ── Helpers de fecha/nombre ───────────────────────────────────

def ultimo_dia(anio: int, mes: int) -> str:
    return date(anio, mes, monthrange(anio, mes)[1]).strftime("%d/%m/%Y")

def num_doc(anio: int, mes: int) -> str:
    return f"{NOMBRES_MES_CORTO[mes]}{anio}"

def nombre_mes(mes: int) -> str:
    return NOMBRES_MES_LARGO[mes]


# ── Resultado de la construccion ──────────────────────────────

@dataclass
class ResultadoPlano:
    filas: list[list]
    recaudo: float
    gastos: float
    db_banco: float
    propietarios_total: float
    recaudos_pend: float
    diferencia_calc: float
    traslado_neto: float
    n_propietarios: int
    advertencias: list[str] = field(default_factory=list)


# ── Constructor principal ─────────────────────────────────────

def construir_filas(
    df_ext: pd.DataFrame,
    df_prop: pd.DataFrame,
    anio: int,
    mes: int,
    n_comprobante: str = "",
    total_sin_conc: float | None = None,
) -> ResultadoPlano:
    """
    Construye las filas del archivo plano PSL a partir de los
    datos del extracto y del informe de movimientos.

    Args:
        df_ext: DataFrame del extracto bancario (parsear_extracto).
        df_prop: DataFrame de propietarios (leer_propietarios).
        anio: Ano contable.
        mes: Mes contable (1-12).
        n_comprobante: Numero de comprobante PSL (vacio = asignar manualmente).
        total_sin_conc: Total del archivo sin conciliar. Si None, se calcula
                        como diferencia entre recaudo y aportes.

    Returns:
        ResultadoPlano con filas y resumen financiero.
    """
    fecha_libro = ultimo_dia(anio, mes)
    doc         = num_doc(anio, mes)
    mes_nom     = nombre_mes(mes)
    advertencias: list[str] = []

    # ── Totales del extracto ──────────────────────────────────
    gastos_comision = abs(df_ext[df_ext["tipo"] == "COMISION"]["valor"].sum())
    gastos_gmf      = abs(df_ext[df_ext["tipo"] == "GMF"]["valor"].sum())
    gastos_iva      = abs(df_ext[df_ext["tipo"] == "IVA_COMISION"]["valor"].sum())
    total_gastos    = gastos_comision + gastos_gmf + gastos_iva

    recaudo  = df_ext[df_ext["tipo"] == "RECAUDO"]["valor"].sum()
    db_banco = recaudo - total_gastos

    traslado_salidas  = abs(df_ext[(df_ext["tipo"] == "TRASLADO") & (df_ext["valor"] < 0)]["valor"].sum())
    traslado_entradas = df_ext[(df_ext["tipo"] == "TRASLADO") & (df_ext["valor"] > 0)]["valor"].sum()
    traslado_neto     = traslado_salidas - traslado_entradas

    # ── Comentarios ───────────────────────────────────────────
    cm = {
        "gastos":   f"GASTOS BANCARIOS DEL MES DE {mes_nom} {anio}",
        "aportes":  f"APORTE MES DE {mes_nom} {anio}",
        "banco":    f"GASTOS BANCARIOS Y APORTES MES DE {mes_nom} {anio}",
        "traslado": f"TRASLADO DE LA CTA CTE A LA FIDUCIARIA BANCOLOMBIA DE {mes_nom} {anio}",
        "pend":     f"RECAUDOS PENDIENTES POR IDENTIFICAR DE {mes_nom} {anio}",
    }

    # ── Funcion auxiliar de fila ──────────────────────────────
    def fila(op, cuenta, tercero, valor, comentario) -> list:
        return [
            COMPANIA, DIVISION, anio, mes, FUENTE, n_comprobante,
            fecha_libro, fecha_libro, op, cuenta, CENTRO, 0,
            tercero, doc, valor, MONEDA, 0, 0, ORIGEN, comentario,
        ]

    filas: list[list] = []

    # ── Gastos bancarios (Debito) ─────────────────────────────
    if gastos_comision > 0:
        filas.append(fila("1", CTA_COMISIONES,   NIT_BANCOLOMBIA, gastos_comision, cm["gastos"]))
    if gastos_gmf > 0:
        filas.append(fila("1", CTA_GMF,           NIT_BANCOLOMBIA, gastos_gmf,      cm["gastos"]))
    if gastos_iva > 0:
        filas.append(fila("1", CTA_IVA_COMISION,  NIT_BANCOLOMBIA, gastos_iva,      cm["gastos"]))

    # ── Aportes por propietario (Credito) ─────────────────────
    for _, row in df_prop.iterrows():
        nit_raw = row["Nro ID Propietario"]
        if pd.isna(nit_raw):
            nit = ""
        else:
            try:
                nit = str(int(float(nit_raw)))
            except (ValueError, TypeError):
                nit = str(nit_raw).strip()
        valor = abs(row["Valor"])
        if nit and valor > 0:
            filas.append(fila("2", CTA_PROPIETARIOS, nit, valor, cm["aportes"]))

    # ── Recaudos pendientes (Credito) ─────────────────────────
    diferencia_calc = recaudo - df_prop["Valor"].sum()
    recaudos_pend   = total_sin_conc if total_sin_conc is not None else diferencia_calc

    if total_sin_conc is not None and abs(total_sin_conc - diferencia_calc) > 1:
        advertencias.append(
            f"El archivo sin conciliar (${total_sin_conc:,.0f}) no coincide "
            f"con la diferencia calculada (${diferencia_calc:,.0f}). "
            f"Diferencia: ${abs(total_sin_conc - diferencia_calc):,.0f}. "
            "Se usa el valor del archivo sin conciliar."
        )

    if recaudos_pend > 0:
        filas.append(fila("2", CTA_PEND_IDENTIF, NIT_BANCOLOMBIA, recaudos_pend, cm["pend"]))

    # ── Debito banco neto ─────────────────────────────────────
    filas.append(fila("1", CTA_BANCO_CTE, "0", db_banco, cm["banco"]))

    # ── Traslado a fiducia ────────────────────────────────────
    if traslado_neto > 0:
        filas.append(fila("2", CTA_BANCO_CTE, "0", traslado_neto, cm["traslado"]))
        filas.append(fila("1", CTA_FIDUCIA,   "0", traslado_neto, cm["traslado"]))

    return ResultadoPlano(
        filas=filas,
        recaudo=recaudo,
        gastos=total_gastos,
        db_banco=db_banco,
        propietarios_total=float(df_prop["Valor"].sum()),
        recaudos_pend=recaudos_pend,
        diferencia_calc=diferencia_calc,
        traslado_neto=traslado_neto,
        n_propietarios=len(df_prop),
        advertencias=advertencias,
    )
