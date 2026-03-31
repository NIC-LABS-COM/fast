"""Funcoes de log do Consumer."""

from datetime import datetime

from .config import LOG_FILE, ALREADY_EXISTS_MARKERS


def log(msg: str) -> None:
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def is_already_exists_error(details: str) -> bool:
    text = (details or "").lower()
    return any(marker in text for marker in ALREADY_EXISTS_MARKERS)
