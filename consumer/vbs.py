"""Resolucao e execucao de scripts VBS locais."""

import os
import subprocess
import tempfile
import traceback

from .config import VBS_DIR
from .logger import log

# Limite seguro para argumentos de linha de comando (Windows max ~8191)
_MAX_ARG_LEN = 7000


def get_vbs_path(filename: str) -> str | None:
    """Retorna o caminho absoluto do VBS dentro da pasta local empacotada."""
    path = os.path.join(VBS_DIR, filename)
    if not os.path.exists(path):
        log(f"VBS nao encontrado: {path}")
        return None
    log(f"VBS local: {path}")
    return path


def execute_vbs(vbs_path: str, args: list[str]) -> tuple[bool, str]:
    if not os.path.exists(vbs_path):
        return False, f"Arquivo nao encontrado: {vbs_path}"
    cscript = os.path.join(os.environ.get("SYSTEMROOT", r"C:\Windows"), "System32", "cscript.exe")

    # Argumentos grandes sao gravados em arquivo temporario.
    # O VBS recebe "@caminho" e le o conteudo do arquivo.
    temp_files = []
    final_args = []
    for arg in args:
        if len(arg) > _MAX_ARG_LEN:
            tf = tempfile.NamedTemporaryFile(
                mode="w", suffix=".txt", delete=False, encoding="utf-8"
            )
            tf.write(arg)
            tf.close()
            temp_files.append(tf.name)
            final_args.append(f"@{tf.name}")
            log(f"Argumento grande ({len(arg)} chars) gravado em {tf.name}")
        else:
            final_args.append(arg)

    cmd = [cscript, "//nologo", vbs_path, *final_args]
    log(f"Executando: {cmd}")
    try:
        result = subprocess.run(cmd, capture_output=True)
    except FileNotFoundError:
        return False, "cscript.exe nao encontrado"
    except Exception as exc:
        return False, str(exc)
    finally:
        for tf_path in temp_files:
            try:
                os.unlink(tf_path)
            except OSError:
                pass

    out = result.stdout.decode("oem", errors="replace").strip() if result.stdout else ""
    err = result.stderr.decode("cp1252", errors="replace").strip() if result.stderr else ""
    log(f"Retorno: code={result.returncode}, stdout='{out}', stderr='{err}'")

    if result.returncode != 0:
        error_msg = err or out or f"Codigo de saida: {result.returncode}"
        # Extrai mensagem limpa de erro SAP
        if error_msg.startswith("SAP_ERROR:"):
            error_msg = error_msg[len("SAP_ERROR:"):].strip()
        return False, error_msg

    if out:
        if err:
            log(f"AVISO: stderr nao vazio (ignorado, returncode=0): {err}")
        return True, out

    if "SAP Frontend Server:" in err:
        return False, err
    return True, "OK"
