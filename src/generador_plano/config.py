# =============================================================
#  config.py — Parametros fijos del proyecto
#  P.A. TOLEDANA DEL SUR (16195) · PSL
# =============================================================
#  Editar solo si cambian las cuentas contables o parametros
#  del fideicomiso en el sistema PSL.
# =============================================================

# ── Identificacion del proyecto ───────────────────────────────
COMPANIA        = "27"
DIVISION        = "27"
CENTRO          = 166
FUENTE          = "0801"
NIT_BANCOLOMBIA = "890903938"
MONEDA          = "PESOC"
ORIGEN          = "Causacion"

# ── Cuentas contables ─────────────────────────────────────────
CTA_COMISIONES   = "53051501"   # Comisiones bancarias
CTA_GMF          = "51159501"   # Gravamen al movimiento financiero (4x1000)
CTA_IVA_COMISION = "53050801"   # IVA sobre comisiones
CTA_PROPIETARIOS = "28051501"   # Aportes propietarios identificados
CTA_PEND_IDENTIF = "28051502"   # Recaudos pendientes por identificar
CTA_FALTANTE     = "28051502"   # Dinero faltante por identificar (ajustar si hay cuenta especifica)
CTA_BANCO_CTE    = "11100607"   # Cuenta corriente banco
CTA_FIDUCIA      = "11253301"   # Cuenta fiduciaria

# ── Contrasenas por defecto (se pueden cambiar en la UI) ──────
PASSWORD_DEFAULT = "901698967"

# ── Encabezados del archivo plano PSL ─────────────────────────
HEADERS = [
    "compania", "division", "ano", "periodo", "fuente", "N comprobante",
    "fecha de contabilizacion", "fecha de la transaccion", "operacion",
    "cod cuenta", "centro", "concepto", "Tercero", "Documento",
    "vlr moneda lc", "cod mond extr", "valor moneda", "valor base",
    "origen", "comentario",
]

# ── Nombres de meses ──────────────────────────────────────────
NOMBRES_MES_CORTO = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC",
}
NOMBRES_MES_LARGO = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
