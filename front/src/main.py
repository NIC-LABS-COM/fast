"""
Ponto de entrada da aplicacao SAP ABAP Dictionary Publisher.

Uso:
    python src/main.py
    python src/main.py --log-level DEBUG

Gerar executavel:
    pyinstaller --onefile src/main.py
"""
import argparse
import os
import sys
import tkinter as tk
from tkinter import ttk

# Garante que 'src/' esteja no sys.path ao rodar diretamente ou via PyInstaller.
_SRC = os.path.dirname(os.path.abspath(__file__))
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from core.logger import setup_logger  # noqa: E402
from ui.app import SapPublisherApp    # noqa: E402


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="SAP ABAP Dictionary Publisher")
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Nivel de log no console (padrao: INFO)",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    logger = setup_logger(log_level=args.log_level)
    logger.info("Iniciando SAP Publisher...")

    root = tk.Tk()
    try:
        ttk.Style(root).theme_use("clam")
    except Exception:
        pass

    SapPublisherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
