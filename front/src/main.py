"""
Ponto de entrada da aplicacao SAP ABAP Dictionary Publisher.

Uso:
    python src/main.py
    python src/main.py --log-level DEBUG
"""
import argparse
import os
import sys

# Garante que 'src/' esteja no sys.path ao rodar diretamente ou via PyInstaller.
_SRC = os.path.dirname(os.path.abspath(__file__))
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from core.logger import setup_logger  # noqa: E402


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

    try:
        import webview
    except ImportError:
        print("\n[ERRO] pywebview nao encontrado.")
        print("Instale com:  pip install pywebview\n")
        sys.exit(1)

    from api import Api

    api = Api()
    html_path = os.path.join(_SRC, "web", "index.html")

    window = webview.create_window(
        "ABAP Dictionary",
        html_path,
        js_api=api,
        width=1450,
        height=920,
        background_color="#111417",
        min_size=(1024, 600),
    )

    def on_start():
        api._set_window(window)

    webview.start(on_start, debug=(args.log_level == "DEBUG"))


if __name__ == "__main__":
    main()
