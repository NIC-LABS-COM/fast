"""Arquitetura legada — queue_vpn_usiminas."""

import json
import traceback
from datetime import datetime

import pika

from .config import QUEUE_RESPONSES
from .logger import log, is_already_exists_error
from .vbs import download_vbs, execute_vbs


def publish_response(channel, action: str, object_name: str, status: str,
                     message: str, correlation_id: str = "") -> None:
    response = {
        "correlationId": correlation_id,
        "action":        action,
        "object_name":   object_name,
        "status":        status,
        "message":       message,
        "timestamp":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    try:
        body = json.dumps(response, ensure_ascii=False)
        channel.basic_publish(
            exchange="", routing_key=QUEUE_RESPONSES,
            body=body.encode("utf-8"),
            properties=pika.BasicProperties(delivery_mode=2),
        )
        log(f"Resposta publicada: {body}")
    except Exception:
        log(f"Erro ao publicar resposta: {traceback.format_exc()}")


def process_legacy_message(channel, body_str: str) -> None:
    """Processa mensagens do formato legado (action / vbs_url / args)."""
    log(f"=== MENSAGEM RECEBIDA === {body_str}")

    try:
        payload = json.loads(body_str)
    except json.JSONDecodeError as e:
        log(f"Erro ao parsear JSON: {e}")
        return

    action         = payload.get("action", "desconhecida")
    vbs_url        = payload.get("vbs_url", "")
    args           = payload.get("args", [])
    object_name    = args[0] if args else "desconhecido"
    correlation_id = payload.get("correlationId", "")

    log(f"Action: {action} | Objeto: {object_name} | CorrelationId: {correlation_id}")
    log(f"VBS URL: {vbs_url}")
    log(f"Args: {args}")

    if not vbs_url:
        log("ERRO: vbs_url vazio.")
        publish_response(channel, action, object_name, "erro",
                         "URL do VBS vazia na mensagem.", correlation_id)
        return

    vbs_path = download_vbs(vbs_url)
    if vbs_path is None:
        publish_response(channel, action, object_name, "erro",
                         "Falha no download do VBS.", correlation_id)
        return

    ok, details = execute_vbs(vbs_path, args)

    if ok:
        log(f"SUCESSO: {action} - {object_name}")
        publish_response(channel, action, object_name, "sucesso", details, correlation_id)
    elif is_already_exists_error(details):
        log(f"AVISO: {action} - {object_name} ja existe.")
        publish_response(channel, action, object_name, "ja_existe",
                         f"{object_name} ja existe no SAP.", correlation_id)
    else:
        log(f"FALHA: {action} - {object_name}: {details}")
        publish_response(channel, action, object_name, "erro", details, correlation_id)


def callback_legacy(ch, _method, _properties, body):
    body_str = body.decode("utf-8")
    try:
        process_legacy_message(ch, body_str)
    except Exception:
        log(f"Erro nao tratado (legado): {traceback.format_exc()}")
