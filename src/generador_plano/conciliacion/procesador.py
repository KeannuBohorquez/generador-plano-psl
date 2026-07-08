# =============================================================
#  conciliacion/procesador.py
#  Logica de negocio para procesar el extracto PDF de la
#  fiduciaria Bancolombia y generar el archivo plano de
#  conciliacion de proyectos (BOSKE, 23LIVING, THECORNER).
# =============================================================

from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber


# ── Patrones regex ─────────────────────────────────────────────

# Capital del Encargo - formato antiguo (encargo + valor en una zona)
_RE_CAPITAL_V1 = re.compile(
    r"(Aporte\s+Traslado Vlr Capital Del Encargo):\s*"
    r'"(\d+)"\s*-\s*(\(?[\d.,\s]+\)?)\s+PNJF'
)

# Capital del Encargo - formato nuevo (valor en linea siguiente)
_RE_CAPITAL_V2 = re.compile(
    r"(Aporte\s+Traslado Vlr Capital Del Encargo):\s*\n"
    r"(\(?[\d.,\s]+\)?)\s+PNJF.*?\n"
    r'"(\d+)"\s*-\s*"',
    re.DOTALL,
)

# Rendimiento del Encargo
_RE_RENDIMIENTO = re.compile(
    r"(Aporte\s+Traslado Vlr Rendimiento Del Encargo):\s*\n"
    r"(\(?[\d.,\s]+\)?)\s+PNJF.*?\n"
    r'"(\d+)"\s*-\s*"',
    re.DOTALL,
)

# Retiro Consig.
_RE_RETIRO_CONSIG = re.compile(
    r"Retiro\s+Consig\.[^-]*-\s*([^-]+?)\s*(?:-\s*)?(?:\d+\s*)?"
    r"\(\s*([\d.,]+)\s*\)\s+PNJF"
)

# Retiro Pago
_RE_RETIRO_PAGO = re.compile(
    r"(Retiro\s+Pago: .*?)\s*-\s*[\d\s.]+\n(\(?[\d.,\s]+\)?)\s+PNJF"
)

# Aporte Sin Fondo
_RE_APORTE_SIN_FONDO = re.compile(
    r"Aporte Aporte\s*:\s*(.*?)\n\s*([\d.,\s]+)\s+PNJF.*?\n"
    r"\d+\s+Aplicar En El Fondo\s+(\d+)",
    re.DOTALL,
)

# Retiro Recaudo Cartera
_RE_RECAUDO = re.compile(
    r"(Retiro Recaudo Cartera Fideicomiso.*?Cliente:)\s*"
    r"\(\s*([\d.,\s]+)\s*\).*?\n(.*?Nro Factura:\s*\d+)",
    re.DOTALL,
)

# Conceptos iniciales
_CONCEPTOS_CLAVE = [
    "Rendimientos después de gastos",
    "Rentención en la fuente",
    "GMF",
    "Costos Transaccionales",
]
_RE_CONCEPTOS = re.compile(
    r"({})".format("|".join(re.escape(c) for c in _CONCEPTOS_CLAVE))
    + r"\s+(\(?\s*[\d.,]+\s*\)?)"
)


# ── Mapeo de proyectos ─────────────────────────────────────────

_PROYECTOS: dict[str, dict] = {
    "BOSKE": {
        "encargos":  ["10010019062-7", "10010021327-4"],
        "compania":  "23",
        "division":  "23",
        "centro":    "164",
    },
    "23LIVING": {
        "encargos":  ["10010019898-0"],
        "compania":  "19",
        "division":  "19",
        "centro":    "227",
    },
    "THECORNER": {
        "encargos":  ["10010019561-1", "10010021806-4"],
        "compania":  "26",
        "division":  "26",
        "centro":    "168",
    },
}

# Nombre de mes en espanol → numero de periodo
_MES_NUM: dict[str, str] = {
    "ENERO": "01", "FEBRERO": "02", "MARZO": "03",     "ABRIL":     "04",
    "MAYO":  "05", "JUNIO":  "06", "JULIO": "07",     "AGOSTO":    "08",
    "SEPTIEMBRE": "09", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12",
}

# NIT Bancolombia (fiduciaria)
_NIT_BANCOLOMBIA = "860531315"


# ── Resultado ─────────────────────────────────────────────────

@dataclass
class ResultadoConciliacion:
    """Resultado del procesamiento del extracto fiduciario."""

    df:            pd.DataFrame  # DataFrame con 20 columnas del archivo plano
    ruta_salida:   str           # Ruta donde se guardo el Excel
    n_filas:       int           # Cantidad de filas generadas
    proyecto:      str           # BOSKE / 23LIVING / THECORNER / DESCONOCIDO
    compania:      str
    division:      str
    centro:        str


# ── Helpers ────────────────────────────────────────────────────

def _parse_valor(s: str) -> float:
    """
    Convierte cadenas como '(1.234,56)' o '1.234,56' a float.
    Los valores entre parentesis se retornan negativos.
    """
    negativo = "(" in s
    limpio = (
        s.replace("(", "").replace(")", "")
         .replace(" ", "").replace(".", "").replace(",", ".")
    )
    val = float(limpio)
    return -val if negativo else val


def _id_a_texto(v) -> str:
    """Convierte un identificador de pandas a texto limpio."""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


# ── Lectura del PDF ───────────────────────────────────────────

def parsear_pdf(ruta: str) -> str:
    """
    Extrae todo el texto del PDF del extracto fiduciario (sin contrasena).

    Args:
        ruta: Ruta al archivo PDF.

    Returns:
        Texto completo del PDF como una sola cadena.
    """
    texto = ""
    with pdfplumber.open(ruta) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto += t + "\n"
    return texto


# ── Deteccion de proyecto ─────────────────────────────────────

def detectar_proyecto(texto: str) -> tuple[str, str, str, str]:
    """
    Detecta el proyecto buscando numeros de encargo en el texto.

    Returns:
        Tupla (nombre_proyecto, compania, division, centro).
        Si no se detecta, devuelve ('DESCONOCIDO', '00', '00', '000').
    """
    for nombre, datos in _PROYECTOS.items():
        for encargo in datos["encargos"]:
            if encargo in texto:
                return nombre, datos["compania"], datos["division"], datos["centro"]
    return "DESCONOCIDO", "00", "00", "000"


# ── Extraccion de transacciones ───────────────────────────────

def _aplicar_patrones(texto: str, mes_nombre: str, anio: str) -> pd.DataFrame:
    """
    Aplica todos los patrones regex al texto del PDF y construye
    un DataFrame con columnas: Numero de Encargo, Valor, Descripcion.

    El orden de filas en el resultado sigue el orden definido en el
    archivo original (conceptos → aportes → retiros).
    """
    filas: list[dict] = []

    # ─ Conceptos iniciales (orden 0) ──────────────────────────
    for m in _RE_CONCEPTOS.finditer(texto):
        filas.append({
            "Numero de Encargo": "",
            "Valor": _parse_valor(m.group(2)),
            "Descripción": f"{m.group(1)} mes de {mes_nombre} de {anio}",
            "_orden": 0,
        })

    # ─ Capital V1 (formato antiguo, orden 1) ──────────────────
    for m in _RE_CAPITAL_V1.finditer(texto):
        filas.append({
            "Numero de Encargo": m.group(2),
            "Valor": _parse_valor(m.group(3)),
            "Descripción": f"{m.group(1)} mes de {mes_nombre} de {anio}",
            "_orden": 1,
        })

    # ─ Rendimiento del encargo (orden 2) ──────────────────────
    for m in _RE_RENDIMIENTO.finditer(texto):
        filas.append({
            "Numero de Encargo": m.group(3),
            "Valor": _parse_valor(m.group(2)),
            "Descripción": f"{m.group(1)} mes de {mes_nombre} de {anio}",
            "_orden": 2,
        })

    # ─ Capital V2 (formato nuevo, orden 3) ────────────────────
    for m in _RE_CAPITAL_V2.finditer(texto):
        filas.append({
            "Numero de Encargo": m.group(3),
            "Valor": _parse_valor(m.group(2)),
            "Descripción": f"{m.group(1)} mes de {mes_nombre} de {anio}",
            "_orden": 3,
        })

    # ─ Aporte Sin Fondo (orden 4) ─────────────────────────────
    for m in _RE_APORTE_SIN_FONDO.finditer(texto):
        val_str = m.group(2).replace(" ", "").replace(".", "").replace(",", ".")
        filas.append({
            "Numero de Encargo": m.group(3),
            "Valor": float(val_str),
            "Descripción": f"Aporte Sin Fondo: {m.group(1).strip()} mes de {mes_nombre} de {anio}",
            "_orden": 4,
        })

    # ─ Retiro Consig. (orden 5) ───────────────────────────────
    for m in _RE_RETIRO_CONSIG.finditer(texto):
        filas.append({
            "Numero de Encargo": "",
            "Valor": -float(m.group(2).replace(".", "").replace(",", ".")),
            "Descripción": f"Retiro Consig. mes de {mes_nombre} de {anio}",
            "_orden": 5,
        })

    # ─ Retiro Pago (orden 6) ──────────────────────────────────
    for m in _RE_RETIRO_PAGO.finditer(texto):
        filas.append({
            "Numero de Encargo": "",
            "Valor": _parse_valor(m.group(2)),
            "Descripción": f"{m.group(1).strip()} mes de {mes_nombre} de {anio}",
            "_orden": 6,
        })

    # ─ Retiro Recaudo Cartera (orden 7) ───────────────────────
    for m in _RE_RECAUDO.finditer(texto):
        val_str = (
            m.group(2).replace("(", "").replace(")", "")
                      .replace(" ", "").replace(".", "").replace(",", ".")
        )
        filas.append({
            "Numero de Encargo": "",
            "Valor": -float(val_str),
            "Descripción": f"{m.group(1).strip()} {m.group(3).strip()}",
            "_orden": 7,
        })

    if not filas:
        return pd.DataFrame(columns=["Numero de Encargo", "Valor", "Descripción"])

    df = pd.DataFrame(filas).sort_values("_orden").drop(columns=["_orden"])
    df["Valor"] = df["Valor"].astype(float)
    return df.reset_index(drop=True)


# ── Calculo de operacion y cuenta ─────────────────────────────

def _op_cuenta(desc: str, valor: float) -> tuple[str, str]:
    """Devuelve (operacion, cod_cuenta) segun la descripcion."""
    d = desc.upper()
    if "RETIRO RECAUDO CARTERA FIDEICOMISO" in d:
        return ("2" if valor < 0 else "1", "Revisar")
    if "RENDIMIENTOS DESPUÉS DE GASTOS" in d:
        return ("2" if valor > 0 else "1", "421005")
    if "RENTENCIÓN EN LA FUENTE" in d:
        return ("2" if valor > 0 else "1", "13551525")
    if "GMF" in d:
        return ("2" if valor > 0 else "1", "511595")
    if "COSTOS TRANSACCIONALES" in d:
        return ("2" if valor > 0 else "1", "530506")
    if "RETIRO CONSIG." in d:
        return ("2" if valor < 0 else "1", "")
    if "APORTE TRASLADO VLR CAPITAL DEL ENCARGO" in d:
        return ("2" if valor > 0 else "1", "28051501")
    if "APORTE TRASLADO VLR RENDIMIENTO DEL ENCARGO" in d:
        return ("2" if valor > 0 else "1", "421006")
    return ("", "Revisar")


# ── Punto de entrada principal ────────────────────────────────

def procesar(
    ruta_pdf:       str,
    ruta_clientes:  str,
    mes_nombre:     str,
    anio:           str,
    ruta_salida:    str,
) -> ResultadoConciliacion:
    """
    Procesa el extracto PDF de la fiduciaria y genera el archivo
    plano de conciliacion en formato Excel de 20 columnas.

    Args:
        ruta_pdf:      Ruta al PDF del extracto (sin contrasena).
        ruta_clientes: Ruta al Excel de clientes (hoja 'CONSOLIDADO ALIANZA').
        mes_nombre:    Nombre del mes en espanol, p.ej. 'MAYO'.
        anio:          Ano como string, p.ej. '2026'.
        ruta_salida:   Ruta donde guardar el archivo Excel resultado.

    Returns:
        ResultadoConciliacion con el DataFrame generado y metadatos.

    Raises:
        ValueError:   Si no se puede leer el PDF o los archivos Excel.
        KeyError:     Si la hoja 'CONSOLIDADO ALIANZA' no existe en clientes.xlsx.
    """
    mes_up = mes_nombre.strip().upper()
    period = _MES_NUM.get(mes_up, "00")
    if period == "00":
        raise ValueError(
            f"Mes no reconocido: '{mes_nombre}'. "
            "Use el nombre completo en espanol (ENERO, FEBRERO, etc.)"
        )

    ultimo_dia = calendar.monthrange(int(anio), int(period))[1]
    fecha_str  = f"{anio}-{period}-{ultimo_dia:02d}"
    fecha_fmt  = datetime.strptime(fecha_str, "%Y-%m-%d").strftime("%d/%m/%Y")

    # 1. Leer PDF
    texto = parsear_pdf(ruta_pdf)
    if not texto.strip():
        raise ValueError("El PDF no contiene texto extraible.")

    # 2. Detectar proyecto
    nombre_proy, compania, division, centro = detectar_proyecto(texto)

    # 3. Extraer transacciones
    df_trans = _aplicar_patrones(texto, mes_nombre.strip().capitalize(), anio)

    # 4. Cruzar con clientes
    df_clientes = pd.read_excel(ruta_clientes, sheet_name="CONSOLIDADO ALIANZA")
    df_clientes.rename(columns={"Encargo": "Numero de Encargo"}, inplace=True)

    df_trans["Numero de Encargo"]    = df_trans["Numero de Encargo"].astype(str)
    df_clientes["Numero de Encargo"] = df_clientes["Numero de Encargo"].astype(str)

    df_cx = pd.merge(df_trans, df_clientes, on="Numero de Encargo", how="left")

    if "Identificación" not in df_cx.columns:
        df_cx["Identificación"] = ""
    if "Cliente" not in df_cx.columns:
        df_cx["Cliente"] = ""

    tercero_col = df_cx["Identificación"].map(_id_a_texto)

    # 5. Construir las 20 columnas del archivo plano
    n = len(df_cx)
    df_plano = pd.DataFrame({
        "compañía":                 [compania]    * n,
        "division":                 [division]    * n,
        "año":                      [anio]         * n,
        "periodo":                  [period]       * n,
        "fuente":                   ["0801"]       * n,
        "N comprobante":            [""]           * n,
        "fecha de contabilizacion": [fecha_fmt]   * n,
        "fecha de la transaccion":  [fecha_fmt]   * n,
        "operación":                [""]           * n,
        "cod cuenta":               [""]           * n,
        "centro":                   [centro]       * n,
        "concepto":                 ["0"]          * n,
        "Tercero":                  tercero_col.values,
        "Documento":                [fecha_fmt]   * n,
        "vlr moneda lc":            df_cx["Valor"].values,
        "cod mond extr":            ["PESOC"]      * n,
        "valor moneda":             [0]            * n,
        "valor base":               [0]            * n,
        "origen":                   ["Causacion"]  * n,
        "comentario":               df_cx["Descripción"].values,
    })

    # 6. Calcular operacion y cuenta contable
    if n > 0:
        ops, cuentas = zip(*[
            _op_cuenta(row["comentario"], row["vlr moneda lc"])
            for _, row in df_plano.iterrows()
        ])
        df_plano["operación"] = list(ops)
        df_plano["cod cuenta"] = list(cuentas)

    # 7. Asignar NIT Bancolombia a conceptos fiduciarios
    mes_cap = mes_nombre.strip().capitalize()
    for concepto in _CONCEPTOS_CLAVE:
        mask = df_plano["comentario"] == f"{concepto} mes de {mes_cap} de {anio}"
        df_plano.loc[mask, "Tercero"] = _NIT_BANCOLOMBIA
    df_plano.loc[df_plano["cod cuenta"] == "421006", "Tercero"] = _NIT_BANCOLOMBIA

    # 8. Guardar
    df_plano.to_excel(ruta_salida, index=False)

    return ResultadoConciliacion(
        df=df_plano,
        ruta_salida=ruta_salida,
        n_filas=n,
        proyecto=nombre_proy,
        compania=compania,
        division=division,
        centro=centro,
    )
