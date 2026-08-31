"""
main.py — Punto de entrada de Data-Travel GUI.

Uso:
    python src/main.py
    python -m src.main
"""
from __future__ import annotations

import sys
from pathlib import Path

# Asegurar que la raiz del proyecto esta en el PATH
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.gui.app_ui import DataTravelApp


def main():
    app = DataTravelApp()
    app.mainloop()


if __name__ == "__main__":
    main()
