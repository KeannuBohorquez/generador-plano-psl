"""
generador_plano
===============
Paquete para la generacion del archivo plano PSL de la
conciliacion bancaria BanColombia — P.A. Toledana del Sur (16195).

Modulos:
    config      — Constantes y cuentas contables
    extractor   — Parseo del extracto bancario PDF
    lector      — Lectura de Excel (propietarios y sin conciliar)
    constructor — Logica de construccion de filas
    exportador  — Exportacion a Excel con formato
    ui.app      — Interfaz grafica tkinter
    ui.widgets  — Componentes de UI reutilizables
"""

__version__ = "1.0.0"
__author__  = "P.A. Toledana del Sur"
