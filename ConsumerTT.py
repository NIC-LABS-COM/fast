import json
import os
import subprocess
import traceback
import urllib.request
import urllib.error
from datetime import datetime

import pika
import ssl

# RabbitMQ
RABBITMQ_HOST = "jackal.rmq.cloudamqp.com"
RABBITMQ_USER = "rhrstugr"
RABBITMQ_PASS = "HC2wvtBtou_DUk9AA276209T4718K9cF"
RABBITMQ_VHOST = "rhrstugr"
QUEUE_NAME = "queue_vpn_usiminas"

# Diretorio temporario para salvar o VBS baixado
TEMP_DIR = os.getenv("TEMP", os.getcwd())
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Log")
os.makedirs(LOG_DIR, exist_ok=True)

LOG_FILE = os.path.join(LOG_DIR, f"consumer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")


def log(msg: str) -> None:
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def download_vbs(url: str) -> str | None:
    """Baixa o VBS do GitHub e retorna o caminho local. None se falhar."""
    # Extrai nome do arquivo da URL
    filename = url.split("/")[-1]
    local_path = os.path.join(TEMP_DIR, filename)

    log(f"Baixando VBS: {url}")
    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=30) as resp:
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
    """Executa o VBS com cscript.exe passando os argumentos."""
    if not os.path.exists(vbs_path):
        log(f"Arquivo VBS nao encontrado: {vbs_path}")
        return False, f"Arquivo nao encontrado: {vbs_path}"

    cmd = ["cscript.exe", "//nologo", vbs_path, *args]
    log(f"Executando: {cmd}")

    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, encoding="cp1252", errors="replace"
        )
    except FileNotFoundError:
        log("cscript.exe nao encontrado no sistema")
        return False, "cscript.exe nao encontrado"
    except Exception as exc:
        log(f"Excecao executando VBS: {exc}")
        return False, str(exc)

    out = (result.stdout or "").strip()
    err = (result.stderr or "").strip()
    log(f"Retorno: code={result.returncode}, stdout='{out}', stderr='{err}'")

    if result.returncode != 0 or "SAP Frontend Server:" in err:
        details = err or out or f"Codigo de saida: {result.returncode}"
        return False, details

    return True, out if out else "OK"


def process_message(body_str: str) -> None:
    """Processa uma mensagem da fila: baixa o VBS e executa."""
    log(f"=== MENSAGEM RECEBIDA === {body_str}")

    try:
        payload = json.loads(body_str)
    except json.JSONDecodeError as e:
        log(f"Erro ao parsear JSON: {e}")
        return

    action = payload.get("action", "desconhecida")
    vbs_url = payload.get("vbs_url", "")
    args = payload.get("args", [])

    log(f"Action: {action}")
    log(f"VBS URL: {vbs_url}")
    log(f"Args: {args}")

    if not vbs_url:
        log("ERRO: vbs_url vazio na mensagem. Ignorando.")
        return

    # Passo 1: Baixar o VBS
    vbs_path = download_vbs(vbs_url)
    if vbs_path is None:
        log("ERRO: Falha no download do VBS. Mensagem nao processada.")
        return

    # Passo 2: Executar o VBS
    ok, details = execute_vbs(vbs_path, args)
    if ok:
        log(f"SUCESSO: {action} executado. Resultado: {details}")
    else:
        log(f"FALHA: {action} falhou. Detalhes: {details}")


def callback(ch, method, properties, body):
    """Callback do RabbitMQ - chamado a cada mensagem."""
    body_str = body.decode("utf-8")
    try:
        process_message(body_str)
    except Exception:
        log(f"Erro nao tratado ao processar mensagem: {traceback.format_exc()}")


def main() -> None:
    log("##################################################")
    log("CONSUMER INICIADO")
    log(f"Fila: {QUEUE_NAME}")
    log(f"Host: {RABBITMQ_HOST}")
    log(f"Log: {LOG_FILE}")
    log("##################################################")

    credentials = pika.PlainCredentials(RABBITMQ_USER, RABBITMQ_PASS)
    context = ssl.create_default_context()
    parameters = pika.ConnectionParameters(
        host=RABBITMQ_HOST,
        port=5671,
        virtual_host=RABBITMQ_VHOST,
        credentials=credentials,
        ssl_options=pika.SSLOptions(context),
    )

    connection = pika.BlockingConnection(parameters)
    channel = connection.channel()
    channel.queue_declare(queue=QUEUE_NAME, durable=True)

    # Processa 1 mensagem por vez (importante para SAP GUI)
    channel.basic_qos(prefetch_count=1)
    channel.basic_consume(queue=QUEUE_NAME, on_message_callback=callback, auto_ack=True)

    log("Aguardando mensagens... (CTRL+C para sair)")
    print("Aguardando mensagens... (CTRL+C para sair)")

    try:
        channel.start_consuming()
    except KeyboardInterrupt:
        log("Consumer encerrado pelo usuario.")
        channel.stop_consuming()
    finally:
        connection.close()
        log("Conexao encerrada.")


if __name__ == "__main__":
    main()