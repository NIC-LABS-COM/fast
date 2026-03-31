"""Download e execucao de scripts VBS."""

import os
import subprocess
import traceback
import urllib.request
import urllib.error

from .config import TEMP_DIR
from .logger import log


def download_vbs(url: str) -> str | None:
    filename   = url.split("/")[-1]
    local_path = os.path.join(TEMP_DIR, filename)
    log(f"Baixando VBS: {url}")
    try:
        with urllib.request.urlopen(urllib.request.Request(url), timeout=30) as resp:
            content = resp.read()
        with open(local_path, "wb") as f:
            f.write(content)
        log(f"VBS salvo em: {local_path} ({len(content)} bytes)")
        return local_path
    except urllib.error.URLError as e:
        log(f"Erro de URL no download: {e}")
        return None
    except Exception:
        log(f"Erro inesperado no download: {traceback.format_exc()}")
        return None


def execute_vbs(vbs_path: str, args: list[str]) -> tuple[bool, str]:
    if not os.path.exists(vbs_path):
        return False, f"Arquivo nao encontrado: {vbs_path}"
    cmd = ["cscript.exe", "//nologo", vbs_path, *args]
    log(f"Executando: {cmd}")
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, encoding="cp1252", errors="replace"
        )
    except FileNotFoundError:
        return False, "cscript.exe nao encontrado"
    except Exception as exc:
        return False, str(exc)

    out = (result.stdout or "").strip()
    err = (result.stderr or "").strip()
    log(f"Retorno: code={result.returncode}, stdout='{out}', stderr='{err}'")

    if result.returncode != 0:
        return False, err or out or f"Codigo de saida: {result.returncode}"

    if out:
        if err:
            log(f"AVISO: stderr nao vazio (ignorado, returncode=0): {err}")
        return True, out

    if "SAP Frontend Server:" in err:
        return False, err
    return True, "OK"
