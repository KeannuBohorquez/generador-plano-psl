# =============================================================
#  extractor.py — Parseo del extracto bancario PDF
# =============================================================

import re
import pdfplumber
import pandas as pd

RE_TRANSACCION = re.compile(
    r"^(\d{1,2}/\d{2})\s+(.+?)\s+(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})$"
)


def clasificar(desc: str) -> str:
    """Clasifica una transaccion por su descripcion."""
    d = desc.upper()
    # Gastos van ANTES de RECAUDO porque algunas descripciones
    # contienen la palabra RECAUDO (ej: "COMIS SERVICIOS DE RECAUDO")
    if "IVA" in d and ("COMIS" in d or "COM REC" in d or "SERVICIO" in d):
        return "IVA_COMISION"
    if "COMIS" in d:               return "COMISION"
    if "IMPTO GOBIERNO" in d:      return "GMF"
    if "TRASLADO FIDUCIARIA" in d: return "TRASLADO"
    if "RECAUDO" in d:             return "RECAUDO"
    # Pagos directos de propietarios/desarrolladores sin la palabra RECAUDO
    # Ej: "PAGO DE PROV ACRECER SAS", "PAGO DE PROV ..."
    if "PAGO DE PROV" in d:        return "RECAUDO"
    return "OTRO"


def parsear_extracto(ruta: str, password: str) -> pd.DataFrame:
    """
    Lee el PDF del extracto bancario BanColombia y retorna un
    DataFrame con columnas: fecha, desc, valor, tipo.

    Args:
        ruta: Ruta al archivo PDF.
        password: Contrasena del PDF.

    Returns:
        DataFrame con las transacciones clasificadas.

    Raises:
        Exception: Si el PDF no se puede abrir o la contrasena es incorrecta.
    """
    rows = []
    with pdfplumber.open(ruta, password=password) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").split("\n"):
                m = RE_TRANSACCION.match(line.strip())
                if m:
                    fecha, desc, val, _ = m.groups()
                    rows.append({
                        "fecha": fecha,
                        "desc":  desc.strip(),
                        "valor": float(val.replace(",", "")),
                    })

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=["fecha", "desc", "valor", "tipo"])
    else:
        df["tipo"] = df["desc"].map(clasificar)
    return df
