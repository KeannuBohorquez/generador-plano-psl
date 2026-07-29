# =============================================================
#  conciliacion/procesador.py
#  Logica de negocio para procesar el extracto PDF de la
#  fiduciaria Bancolombia y generar el archivo plano de
#  conciliacion de proyectos (BOSKE, 23LIVING, THECORNER).
# =============================================================

from __future__ import annotations

import calendar
import io
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import msoffcrypto
import pandas as pd
import pdfplumber
from msoffcrypto.exceptions import InvalidKeyError
from pdfminer.pdfdocument import PDFPasswordIncorrect


class PasswordError(Exception):
    """
    Se lanza cuando un archivo (PDF o Excel) esta protegido con
    contrasena y no se dio ninguna, o la que se dio es incorrecta.
    """


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


def _normalizar_txt(s: str) -> str:
    """
    Normaliza un texto para comparaciones tolerantes:
    mayusculas, sin acentos, espacios internos colapsados y
    sin espacios al inicio/fin.
    """
    s = str(s).strip().upper()
    for a, b in (
        ("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U"), ("Ñ", "N"),
    ):
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s)


# ── Lectura del archivo de clientes (formato base ALIANZA) ────

# Nombre esperado de la hoja de clientes (tolerante a espacios/mayus)
_HOJA_CLIENTES = "CONSOLIDADO ALIANZA"

# Alias aceptados para cada columna estandar (normalizados)
_ALIAS_ENCARGO = {"ENCARGO", "NUMERO DE ENCARGO", "N ENCARGO", "NO ENCARGO"}
_ALIAS_IDENT   = {"IDENTIFICACION", "IDENTIFICACION CLIENTE", "NIT", "CEDULA"}
_ALIAS_CLIENTE = {"CLIENTE", "NOMBRE CLIENTE", "NOMBRE"}


def _buscar_hoja_clientes(ruta: str) -> str:
    """
    Busca en el archivo Excel la hoja de clientes tolerando
    diferencias de mayusculas/espacios (p.ej. 'CONSOLIDADO ALIANZA ').

    Args:
        fuente: Ruta al archivo (str) o buffer ya desencriptado
                (io.BytesIO), tal como lo retorna _abrir_fuente_excel().

    Raises:
        ValueError: Si no se encuentra ninguna hoja compatible.
    """
    xls = pd.ExcelFile(fuente)
    for nombre in xls.sheet_names:
        if _normalizar_txt(nombre) == _HOJA_CLIENTES:
            return nombre
    raise ValueError(
        f"No se encontro la hoja 'CONSOLIDADO ALIANZA'.\n"
        f"Hojas disponibles: {', '.join(xls.sheet_names)}"
    )


def _abrir_fuente_excel(ruta: str, pwd: str = "") -> str | io.BytesIO:
    """
    Devuelve una fuente legible por pandas para el archivo de clientes:
    la ruta original si no esta protegido, o un buffer ya desencriptado
    si lo esta.

    Args:
        ruta: Ruta al archivo Excel de clientes/cartera.
        pwd: Contrasena, si el archivo esta protegido.

    Returns:
        La misma ruta (str) si el archivo no esta protegido, o un
        io.BytesIO con el contenido ya desencriptado.

    Raises:
        PasswordError: Si el archivo esta protegido y no se dio
                       contrasena, o la contrasena es incorrecta.
    """
    try:
        with open(ruta, "rb") as f:
            of = msoffcrypto.OfficeFile(f)
            if not of.is_encrypted():
                return ruta
            if not pwd:
                raise PasswordError(
                    f"El archivo de clientes requiere contrasena: "
                    f"{Path(ruta).name}"
                )
            try:
                of.load_key(password=pwd)
                buf = io.BytesIO()
                of.decrypt(buf)
            except InvalidKeyError as ex:
                raise PasswordError(
                    f"La contrasena del archivo de clientes es "
                    f"incorrecta: {Path(ruta).name}"
                ) from ex
            buf.seek(0)
            return buf
    except PasswordError:
        raise
    except Exception:
        # msoffcrypto no pudo interpretar el archivo (no es OLE / no
        # esta encriptado de forma estandar) -> se asume legible directo
        return ruta


def _leer_clientes(ruta: str, pwd: str = "") -> pd.DataFrame:
    """
    Lee el archivo de clientes en su formato base (el mismo que se
    usa en la hoja 'CONSOLIDADO ALIANZA' del archivo de cartera,
    p.ej. 'Cartera_BOSKE APTOS_Mayo 31_2026.xlsx') y normaliza sus
    columnas a: 'Numero de Encargo', 'Identificación', 'Cliente'.

    Tolera variaciones de mayusculas/espacios tanto en el nombre de
    la hoja como en los encabezados de columna (p.ej. 'ENCARGO ',
    'IDENTIFICACIÓN', 'CLIENTE'), y soporta archivos protegidos con
    contrasena.

    Args:
        ruta: Ruta al archivo Excel de clientes/cartera.
        pwd: Contrasena del archivo, si esta protegido.

    Returns:
        DataFrame con columnas 'Numero de Encargo' (texto, sin
        decimales), 'Identificación' y 'Cliente'. Las filas sin
        numero de encargo se descartan.

    Raises:
        ValueError: Si no se encuentra la hoja o la columna de encargo.
        PasswordError: Si el archivo esta protegido y no se dio
                       contrasena, o la contrasena es incorrecta.
    """
    fuente = _abrir_fuente_excel(ruta, pwd)
    hoja = _buscar_hoja_clientes(fuente)
    if hasattr(fuente, "seek"):
        fuente.seek(0)
    df = pd.read_excel(fuente, sheet_name=hoja)

    # Mapear columnas reales -> nombres normalizados
    col_encargo = col_ident = col_cliente = None
    for col in df.columns:
        norm = _normalizar_txt(col)
        if col_encargo is None and norm in _ALIAS_ENCARGO:
            col_encargo = col
        elif col_ident is None and norm in _ALIAS_IDENT:
            col_ident = col
        elif col_cliente is None and norm in _ALIAS_CLIENTE:
            col_cliente = col

    if col_encargo is None:
        raise ValueError(
            f"No se encontro la columna 'Encargo' en la hoja '{hoja}'.\n"
            f"Columnas disponibles: {', '.join(str(c) for c in df.columns)}"
        )

    renombres = {col_encargo: "Numero de Encargo"}
    if col_ident is not None:
        renombres[col_ident] = "Identificación"
    if col_cliente is not None:
        renombres[col_cliente] = "Cliente"
    df = df.rename(columns=renombres)

    if "Identificación" not in df.columns:
        df["Identificación"] = ""
    if "Cliente" not in df.columns:
        df["Cliente"] = ""

    # Descartar filas sin numero de encargo (filas vacias al final)
    df = df[df["Numero de Encargo"].notna()].copy()

    # Normalizar encargo e identificacion a texto limpio (sin '.0')
    df["Numero de Encargo"] = df["Numero de Encargo"].map(_id_a_texto)
    df["Identificación"]    = df["Identificación"].map(_id_a_texto)

    return df[["Numero de Encargo", "Identificación", "Cliente"]]


# ── Lectura del PDF ───────────────────────────────────────────

def parsear_pdf(ruta: str, password: str = "") -> str:
    """
    Extrae todo el texto del PDF del extracto fiduciario. Soporta
    PDFs protegidos con contrasena (la mayoria no lo estan, pero
    algunos extractos si vienen protegidos).

    Args:
        ruta: Ruta al archivo PDF.
        password: Contrasena del PDF, si aplica. Cadena vacia si
                  el PDF no esta protegido.

    Returns:
        Texto completo del PDF como una sola cadena.

    Raises:
        PasswordError: Si el PDF esta protegido y no se dio
                       contrasena, o la contrasena es incorrecta.
    """
    texto = ""
    try:
        with pdfplumber.open(ruta, password=password or "") as pdf:
            for pagina in pdf.pages:
                t = pagina.extract_text()
                if t:
                    texto += t + "\n"
    except Exception as ex:
        if _es_password_incorrecta_pdf(ex):
            raise PasswordError(
                f"El extracto PDF requiere contrasena o la contrasena "
                f"es incorrecta: {Path(ruta).name}"
            ) from ex
        raise
    return texto


def _es_password_incorrecta_pdf(ex: Exception) -> bool:
    """Detecta si una excepcion de pdfplumber/pdfminer es por contrasena."""
    if isinstance(ex, PDFPasswordIncorrect):
        return True
    # pdfplumber envuelve la excepcion original de pdfminer en
    # PdfminerException, guardandola en args[0] y/o __context__.
    origen = ex.args[0] if ex.args else None
    if isinstance(origen, PDFPasswordIncorrect):
        return True
    if isinstance(getattr(ex, "__context__", None), PDFPasswordIncorrect):
        return True
    return False


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


# ── Verificacion previa de contrasenas ─────────────────────────

def verificar_acceso(ruta_pdf: str, ruta_clientes: str, password: str = "") -> None:
    """
    Verifica que se puedan abrir el PDF y el Excel de clientes con
    la contrasena dada, sin procesar todo el contenido. Pensado para
    validar en la UI (hilo principal) ANTES de lanzar el procesamiento
    pesado en un hilo de fondo, para poder pedir la contrasena con un
    dialogo si hace falta.

    Args:
        ruta_pdf:      Ruta al PDF del extracto.
        ruta_clientes: Ruta al Excel de clientes/cartera.
        password:      Contrasena a probar en ambos archivos (puede
                        ser vacia si ninguno esta protegido).

    Raises:
        PasswordError: Si alguno de los dos archivos esta protegido
                       y la contrasena dada esta vacia o es incorrecta.
    """
    # PDF: abrir y tocar la primera pagina es suficiente para que
    # pdfminer valide la contrasena (no hace falta extraer texto).
    try:
        with pdfplumber.open(ruta_pdf, password=password or "") as pdf:
            _ = pdf.pages[0] if pdf.pages else None
    except Exception as ex:
        if _es_password_incorrecta_pdf(ex):
            raise PasswordError(
                f"El extracto PDF requiere contrasena o la contrasena "
                f"es incorrecta: {Path(ruta_pdf).name}"
            ) from ex
        raise

    # Excel de clientes: _abrir_fuente_excel ya valida la contrasena.
    _abrir_fuente_excel(ruta_clientes, password)


# ── Punto de entrada principal ────────────────────────────────

def procesar(
    ruta_pdf:       str,
    ruta_clientes:  str,
    mes_nombre:     str,
    anio:           str,
    ruta_salida:    str,
    password:       str = "",
) -> ResultadoConciliacion:
    """
    Procesa el extracto PDF de la fiduciaria y genera el archivo
    plano de conciliacion en formato Excel de 20 columnas.

    Args:
        ruta_pdf:      Ruta al PDF del extracto.
        ruta_clientes: Ruta al Excel de clientes/cartera. Acepta el
                       formato base (p.ej. 'Cartera_BOSKE APTOS_...xlsx')
                       con la hoja 'CONSOLIDADO ALIANZA' (tolerante a
                       espacios/mayusculas en hoja y columnas: 'ENCARGO ',
                       'IDENTIFICACIÓN', 'CLIENTE', etc.).
        mes_nombre:    Nombre del mes en espanol, p.ej. 'MAYO'.
        anio:          Ano como string, p.ej. '2026'.
        ruta_salida:   Ruta donde guardar el archivo resultado (.xls).
        password:      Contrasena compartida para el PDF y/o el Excel
                       de clientes, si alguno esta protegido. Cadena
                       vacia si ninguno lo esta.

    Returns:
        ResultadoConciliacion con el DataFrame generado y metadatos.

    Raises:
        ValueError:    Si no se puede leer el PDF, o si el Excel de
                       clientes no tiene la hoja/columna de encargo.
        PasswordError: Si el PDF o el Excel de clientes estan
                       protegidos y la contrasena esta vacia o es
                       incorrecta.
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
    texto = parsear_pdf(ruta_pdf, password)
    if not texto.strip():
        raise ValueError("El PDF no contiene texto extraible.")

    # 2. Detectar proyecto
    nombre_proy, compania, division, centro = detectar_proyecto(texto)

    # 3. Extraer transacciones
    df_trans = _aplicar_patrones(texto, mes_nombre.strip().capitalize(), anio)

    # 4. Cruzar con clientes (formato base ALIANZA, tolerante a
    #    variaciones de mayusculas/espacios en hoja y columnas)
    df_clientes = _leer_clientes(ruta_clientes, password)

    df_trans["Numero de Encargo"] = df_trans["Numero de Encargo"].astype(str)
    # df_clientes["Numero de Encargo"] ya viene normalizado por _leer_clientes

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

    # 8. Guardar en formato .xls (Excel 97-2003), con los mismos
    #    formatos de columna (texto/numero/general) que usa Toledana
    from ..exportador import exportar_plano_xls
    exportar_plano_xls(
        headers=df_plano.columns.tolist(),
        filas=df_plano.values.tolist(),
        ruta_salida=ruta_salida,
    )

    return ResultadoConciliacion(
        df=df_plano,
        ruta_salida=ruta_salida,
        n_filas=n,
        proyecto=nombre_proy,
        compania=compania,
        division=division,
        centro=centro,
    )
