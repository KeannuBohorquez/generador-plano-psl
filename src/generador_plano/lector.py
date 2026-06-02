# =============================================================
#  lector.py — Lectura de archivos Excel
#  - Informe de movimientos por propietario (protegido)
#  - Movimientos sin conciliar
# =============================================================

import io

import msoffcrypto
import pandas as pd


# ── Helpers ───────────────────────────────────────────────────

def _desencriptar(ruta: str, pwd: str) -> io.BytesIO:
    """
    Desencripta un Excel protegido con msoffcrypto.
    Si el archivo no tiene contrasena, lo devuelve como stream directo.
    """
    try:
        with open(ruta, "rb") as f:
            of = msoffcrypto.OfficeFile(f)
            if not of.is_encrypted():
                return open(ruta, "rb")
            of.load_key(password=pwd)
            buf = io.BytesIO()
            of.decrypt(buf)
        buf.seek(0)
        return buf
    except Exception:
        return open(ruta, "rb")


def _leer_excel(ruta: str, pwd: str | None = None) -> pd.DataFrame:
    """
    Lee un Excel en cualquier formato soportado:
      - .xlsx  (openpyxl)
      - .xls   (xlrd)
      - .xlsx encriptado (msoffcrypto → openpyxl)
      - .xls encriptado / OLE2 (msoffcrypto → xlrd)

    Raises:
        ValueError: Si ninguna estrategia logra abrir el archivo.
    """
    # 1. Intentar desencriptar con msoffcrypto
    if pwd:
        try:
            with open(ruta, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                if of.is_encrypted():
                    of.load_key(password=pwd)
                    buf = io.BytesIO()
                    of.decrypt(buf)
                    buf.seek(0)
                    return pd.read_excel(buf, header=0, engine="openpyxl")
        except ValueError:
            raise
        except Exception:
            pass

    # 2. Intentar openpyxl directo (.xlsx sin contrasena)
    try:
        return pd.read_excel(ruta, header=0, engine="openpyxl")
    except Exception:
        pass

    # 3. Fallback xlrd (.xls o formatos antiguos)
    try:
        return pd.read_excel(ruta, header=0, engine="xlrd")
    except Exception as e:
        raise ValueError(
            f"No se pudo leer el archivo '{ruta}'.\n"
            f"Detalle: {e}\n"
            "Verifica que no este abierto en Excel y sea un formato valido."
        ) from e


def _normalizar(serie: pd.Series) -> pd.Series:
    """Elimina acentos y convierte a mayusculas para comparaciones robustas."""
    return (
        serie.astype(str)
        .str.upper()
        .str.replace("Á", "A", regex=False)
        .str.replace("á", "A", regex=False)
        .str.replace("É", "E", regex=False)
        .str.replace("é", "E", regex=False)
        .str.replace("Í", "I", regex=False)
        .str.replace("í", "I", regex=False)
        .str.replace("Ó", "O", regex=False)
        .str.replace("ó", "O", regex=False)
        .str.replace("Ú", "U", regex=False)
        .str.replace("ú", "U", regex=False)
        .str.replace("Ñ", "N", regex=False)
        .str.replace("ñ", "N", regex=False)
    )


# ── Funciones publicas ────────────────────────────────────────

def leer_propietarios(
    ruta: str, password: str, anio: int, mes: int
) -> pd.DataFrame:
    """
    Lee la hoja 'Mov_Por_Propietario' del informe de movimientos y
    filtra por mes usando la columna 'Fecha Mov. Banco' (no Fecha Contable).

    Args:
        ruta: Ruta al Excel protegido.
        password: Contrasena del Excel.
        anio: Ano a filtrar.
        mes: Mes a filtrar (1-12).

    Returns:
        DataFrame con los aportes del mes por propietario.
    """
    buf = _desencriptar(ruta, password)
    df  = pd.read_excel(buf, sheet_name="Mov_Por_Propietario",
                        header=6, engine="openpyxl")

    df["Fecha Mov. Banco"] = pd.to_datetime(
        df["Fecha Mov. Banco"], errors="coerce"
    )
    mask = (
        (df["Fecha Mov. Banco"].dt.year  == anio) &
        (df["Fecha Mov. Banco"].dt.month == mes)  &
        _normalizar(df["Tipo Movimiento"]).str.contains(
            "APORTES CUENTAS BANCARIAS", na=False
        )
    )
    return df[mask].copy()


def leer_sin_conciliar(
    ruta: str, anio: int, mes: int, pwd: str | None = None
) -> tuple[float, int]:
    """
    Lee el archivo de movimientos sin conciliar y retorna el total
    del mes indicado y el numero de registros.

    Args:
        ruta: Ruta al archivo Excel (puede ser .xls o .xlsx, con o sin password).
        anio: Ano a filtrar.
        mes: Mes a filtrar (1-12).
        pwd: Contrasena opcional.

    Returns:
        Tupla (total_mes, n_registros).
    """
    df = _leer_excel(ruta, pwd=pwd)

    # Buscar columna de fecha: preferir "Fecha Banco"
    date_col = None
    for col in df.columns:
        if "fecha" in str(col).lower() and "banco" in str(col).lower():
            date_col = col
            break
    if date_col is None:
        for col in df.columns:
            if "fecha" in str(col).lower():
                date_col = col
                break

    if date_col:
        df[date_col] = pd.to_datetime(
            df[date_col], errors="coerce", dayfirst=True
        )
        df = df[
            (df[date_col].dt.year == anio) &
            (df[date_col].dt.month == mes)
        ]

    if "Valor" not in df.columns:
        return 0.0, 0

    val = df["Valor"]
    if val.dtype == object:
        val = (
            val.astype(str)
            .str.replace(r"[$,\s]", "", regex=True)
        )
        val = pd.to_numeric(val, errors="coerce").fillna(0)

    return float(val.sum()), len(df)
