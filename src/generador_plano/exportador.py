# =============================================================
#  exportador.py — Exportacion del archivo plano a Excel (.xls)
#
#  Todo archivo generado por la aplicacion (Toledana y
#  Conciliacion Proyectos) se guarda en formato binario
#  Excel 97-2003 (.xls) usando xlwt, ya que PSL requiere ese
#  formato para importar. openpyxl/pandas.to_excel no pueden
#  escribir .xls, por eso se usa xlwt aqui.
# =============================================================

from __future__ import annotations

import xlwt

from .config import HEADERS

# Anchos de columna en caracteres (mismo orden que HEADERS)
_ANCHOS = [10, 10, 6, 8, 8, 14, 22, 22, 10, 12, 8, 10, 14, 10, 16, 12, 12, 12, 12, 50]

# Colores personalizados (indices de paleta xlwt, rango libre 0x08-0x3F)
_IDX_COLOR_HEADER = 0x21   # azul header  (1F4E79)
_IDX_COLOR_ALT    = 0x22   # azul claro alterno (DCE6F1)
_RGB_HEADER = (0x1F, 0x4E, 0x79)
_RGB_ALT    = (0xDC, 0xE6, 0xF1)

# Formato de celda por columna (indice 1-based).
# Columnas marcadas como texto ("@") reciben ademas conversion a str
# para evitar que Excel interprete numeros largos (NITs, cuentas) en
# notacion cientifica o pierda ceros a la izquierda.
#
# Mapa: col_idx -> (number_format, forzar_texto)
#   number_format : cadena de formato de Excel (num_format_str de xlwt)
#   forzar_texto  : si True, el valor se convierte a str antes de escribir
#
# Orden de columnas (identico en Toledana y Conciliacion Proyectos,
# ambos generan el mismo archivo plano PSL de 20 columnas):
#  1 compania | 2 division | 3 ano | 4 periodo | 5 fuente
#  6 N comprobante | 7 fecha contab. | 8 fecha transac. | 9 operacion
# 10 cod cuenta | 11 centro | 12 concepto | 13 Tercero | 14 Documento
# 15 vlr moneda lc | 16 cod mond extr | 17 valor moneda | 18 valor base
# 19 origen | 20 comentario
_COL_FMT: dict[int, tuple[str, bool]] = {
    1:  ("General", False),  # compania
    2:  ("General", False),  # division
    3:  ("General", False),  # ano
    4:  ("General", False),  # periodo
    5:  ("@",       True ),  # fuente            → texto
    6:  ("@",       True ),  # N comprobante     → texto
    7:  ("@",       True ),  # fecha contab.     → texto dd/mm/yyyy
    8:  ("@",       True ),  # fecha transac.    → texto dd/mm/yyyy
    9:  ("@",       True ),  # operacion         → texto
    10: ("@",       True ),  # cod cuenta        → texto
    11: ("General", False),  # centro            → general
    12: ("General", False),  # concepto          → general
    13: ("@",       True ),  # Tercero           → texto
    14: ("@",       True ),  # Documento         → texto
    15: ("#,##0.00",False),  # vlr moneda lc     → Numero
    16: ("@",       True ),  # cod mond extr     → texto
    17: ("General", False),  # valor moneda      → general
    18: ("General", False),  # valor base        → general
    19: ("@",       True ),  # origen            → texto
    20: ("General", False),  # comentario        → general
}


def _str_seguro(val) -> str:
    """Convierte un valor a str, devolviendo '' si es None."""
    if val is None:
        return ""
    return str(val)


def _construir_estilos(wb: xlwt.Workbook) -> tuple[dict, dict, dict, "xlwt.XFStyle"]:
    """
    Registra los colores personalizados en la paleta del workbook y
    arma los estilos (xlwt.XFStyle) por columna para tres estados:
    normal, fila alterna y fila de dinero faltante.

    Returns:
        Tupla (estilos_normal, estilos_alt, estilos_faltante, header_style),
        cada uno de los primeros tres es un dict {col_idx_1based: XFStyle}.
    """
    wb.set_colour_RGB(_IDX_COLOR_HEADER, *_RGB_HEADER)
    wb.set_colour_RGB(_IDX_COLOR_ALT,    *_RGB_ALT)

    header_style = xlwt.easyxf(
        f"pattern: pattern solid, fore_colour {_IDX_COLOR_HEADER};"
        "font: colour white, bold on, height 200;"
        "align: horiz center;"
    )

    estilos_normal:    dict[int, xlwt.XFStyle] = {}
    estilos_alt:       dict[int, xlwt.XFStyle] = {}
    estilos_faltante:  dict[int, xlwt.XFStyle] = {}

    for ci, (fmt, _) in _COL_FMT.items():
        estilos_normal[ci] = xlwt.easyxf(num_format_str=fmt)
        estilos_alt[ci] = xlwt.easyxf(
            f"pattern: pattern solid, fore_colour {_IDX_COLOR_ALT};",
            num_format_str=fmt,
        )
        estilos_faltante[ci] = xlwt.easyxf(
            "pattern: pattern solid, fore_colour red;"
            "font: colour white, bold on;",
            num_format_str=fmt,
        )

    return estilos_normal, estilos_alt, estilos_faltante, header_style


def exportar_plano_xls(
    headers: list[str],
    filas: list[list],
    ruta_salida: str,
    indices_faltante: list[int] | None = None,
    nombre_hoja: str = "Archivo Plano",
) -> None:
    """
    Escribe cualquier archivo plano PSL de 20 columnas en formato
    .xls (Excel 97-2003), aplicando los mismos formatos de columna
    (texto/numero/general) sin importar el modulo que lo genere
    (Toledana o Conciliacion Proyectos).

    Args:
        headers: Lista de encabezados (deben mantener el mismo orden
                 posicional que _COL_FMT: compania, division, ano...).
        filas: Lista de filas (listas de 20 valores cada una).
        ruta_salida: Ruta completa del archivo .xls a crear.
        indices_faltante: Indices base-0 de filas que representan
                          dinero faltante (se resaltan en rojo).
        nombre_hoja: Nombre de la hoja dentro del libro.

    Raises:
        PermissionError: Si el archivo esta abierto en Excel.
    """
    indices_faltante = set(indices_faltante or [])

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(nombre_hoja)

    estilos_normal, estilos_alt, estilos_faltante, header_style = _construir_estilos(wb)

    # ── Encabezado (fila 0 en xlwt, 0-based) ──────────────────
    for ci, h in enumerate(headers):
        ws.write(0, ci, h, header_style)

    # ── Datos con filas alternas ──────────────────────────────
    for ri, fila in enumerate(filas, start=1):
        idx_base0        = ri - 1
        es_faltante       = idx_base0 in indices_faltante
        excel_row_1based  = ri + 1   # equivalente a la fila visible en Excel (header=1)
        es_alt            = (excel_row_1based % 2 == 0)

        for ci, val in enumerate(fila):
            col_idx = ci + 1
            fmt, forzar_texto = _COL_FMT.get(col_idx, ("General", False))

            if forzar_texto:
                val = _str_seguro(val)

            if es_faltante:
                style = estilos_faltante.get(col_idx, estilos_faltante[1])
            elif es_alt:
                style = estilos_alt.get(col_idx, estilos_alt[1])
            else:
                style = estilos_normal.get(col_idx, estilos_normal[1])

            ws.write(ri, ci, val, style)

    # ── Anchos y panel fijo ────────────────────────────────────
    for ci, ancho in enumerate(_ANCHOS):
        if ci < len(headers):
            ws.col(ci).width = int(256 * (ancho + 2))

    ws.panes_frozen  = True
    ws.horz_split_pos = 1

    try:
        wb.save(ruta_salida)
    except OSError as e:
        # xlwt no distingue PermissionError de otros OSError en Windows
        # cuando el archivo esta abierto en Excel; se propaga tal cual
        # para que la capa de UI lo detecte.
        raise PermissionError(str(e)) from e


def exportar_excel(
    filas: list[list],
    ruta_salida: str,
    indices_faltante: list[int] | None = None,
) -> None:
    """
    Escribe las filas del archivo plano PSL (Toledana del Sur) en
    un .xls con formato.

    Formatos de columna aplicados segun requerimiento:
    - Texto (@)  : fuente, N comprobante, fechas, operacion, cod cuenta,
                   Tercero, Documento, cod mond extr, origen
    - Numero     : vlr moneda lc  (#,##0.00)
    - General    : compania, division, ano, periodo, centro, concepto,
                   valor moneda, valor base, comentario

    Args:
        filas: Lista de filas generadas por construir_filas().
        ruta_salida: Ruta completa del archivo .xls a crear.
        indices_faltante: Indices base-0 de filas que representan
                          dinero faltante (se resaltan en rojo).

    Raises:
        PermissionError: Si el archivo esta abierto en Excel.
    """
    exportar_plano_xls(HEADERS, filas, ruta_salida, indices_faltante)
