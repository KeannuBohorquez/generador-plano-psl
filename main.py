"""
main.py — Punto de entrada del Generador Archivo Plano PSL
==========================================================
Ejecutar directamente:
    python main.py

O compilar a .exe con scripts/COMPILAR.bat
"""

import sys
import traceback

# Agregar src/ al path para que el paquete sea localizable
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from tkinter import messagebox


def main() -> None:
    from generador_plano.ui.app import App
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        tb = traceback.format_exc()
        try:
            messagebox.showerror(
                "Error critico",
                f"No se pudo iniciar la aplicacion:\n\n{e}\n\n{tb}",
            )
        except Exception:
            print(f"Error critico:\n{e}\n{tb}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
