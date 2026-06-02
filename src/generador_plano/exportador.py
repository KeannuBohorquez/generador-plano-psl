# =============================================================
#  exportador.py — Exportacion del archivo plano a Excel
# =============================================================

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .config import HEADERS

# Anchos de columna en caracteres (mismo orden que HEADERS)
_ANCHOS = [10, 10, 6, 8, 8, 14, 22, 22, 10, 12, 8, 10, 14, 10, 16, 12, 12, 12, 12, 50]


def exportar_excel(
    filas: list[list],
    ruta_salida: str,
    indices_faltante: list[int] | None = None,
) -> None:
    """
    Escribe las filas del archivo plano en un Excel con formato.

    Args:
        filas: Lista de filas generadas por construir_filas().
        ruta_salida: Ruta completa del archivo .xlsx a crear.
        indices_faltante: Indices base-0 de filas que representan
                          dinero faltante (se resaltan en rojo).

    Raises:
        PermissionError: Si el archivo esta abierto en Excel.
    """
    indices_faltante = set(indices_faltante or [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Archivo Plano"

    # ── Encabezado ────────────────────────────────────────────
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # ── Datos con filas alternas ──────────────────────────────
    alt_fill      = PatternFill("solid", fgColor="DCE6F1")
    faltante_fill = PatternFill("solid", fgColor="FF0000")   # rojo fuerte
    faltante_font = Font(color="FFFFFF", bold=True, size=10)

    for ri, fila in enumerate(filas, 2):
        idx_base0 = ri - 2   # indice base-0 de la fila
        es_faltante = idx_base0 in indices_faltante

        if es_faltante:
            fill = faltante_fill
        elif ri % 2 == 0:
            fill = alt_fill
        else:
            fill = None

        for ci, val in enumerate(fila, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            if fill:
                cell.fill = fill
            if es_faltante:
                cell.font = faltante_font
            if ci == 15:  # vlr moneda lc
                cell.number_format = "#,##0.00"

    # ── Anchos y panel fijo ───────────────────────────────────
    for ci, ancho in enumerate(_ANCHOS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = ancho
    ws.freeze_panes = "A2"

    wb.save(ruta_salida)
