"""Nova arquitetura — q.usiminas.v1 / routing keys."""

import json
import traceback

import pika

from .config import (
    QUEUE_RESPONSES, VBS_BY_ROUTING_KEY, QUERY_ROUTING_KEYS, ROUTING_KEY_PREFIX,
)
from .logger import log
from .parsers import parse_requests_txt, parse_reports_txt, parse_packages_txt, parse_versions_metadata_txt, parse_abap_files_by_request_txt
from .vbs import download_vbs, execute_vbs


# ------------------------------------------------------------------ #
#  Publicacao de respostas
# ------------------------------------------------------------------ #
def publish_fix_response(channel, reply_to: str, has_error: bool,
                         error_message: str, is_code_error: bool,
                         correlation_id: str = "") -> None:
    """Publica resposta no formato novo: {hasError, errorMessage, isCodeError}."""
    response = {
        "hasError":      has_error,
        "errorMessage":  error_message if has_error else None,
        "isCodeError":   is_code_error,
        "correlationId": correlation_id,
    }
    try:
        body = json.dumps(response, ensure_ascii=False)
        props = pika.BasicProperties(
            delivery_mode=2,
            correlation_id=correlation_id if correlation_id else None,
            content_type="application/json",
        )
        channel.basic_publish(
            exchange="", routing_key=reply_to,
            body=body.encode("utf-8"),
            properties=props,
        )
        log(f"Resposta FIX publicada em '{reply_to}': {body}")
    except Exception:
        log(f"Erro ao publicar resposta FIX: {traceback.format_exc()}")


def publish_query_response(channel, reply_to: str, data: list,
                           correlation_id: str = "") -> None:
    """Publica resposta de query: lista pura de objetos."""
    try:
        body = json.dumps(data, ensure_ascii=False)
        props = pika.BasicProperties(
            delivery_mode=2,
            correlation_id=correlation_id if correlation_id else None,
            content_type="application/json",
        )
        channel.basic_publish(
            exchange="", routing_key=reply_to,
            body=body.encode("utf-8"),
            properties=props,
        )
        log(f"Resposta QUERY publicada em '{reply_to}': {body}")
    except Exception:
        log(f"Erro ao publicar resposta QUERY: {traceback.format_exc()}")


def publish_string_response(channel, reply_to: str, data: str,
                            correlation_id: str = "") -> None:
    """Publica resposta de query que retorna string simples."""
    try:
        body = json.dumps(data, ensure_ascii=False)
        props = pika.BasicProperties(
            delivery_mode=2,
            correlation_id=correlation_id if correlation_id else None,
            content_type="application/json",
        )
        channel.basic_publish(
            exchange="", routing_key=reply_to,
            body=body.encode("utf-8"),
            properties=props,
        )
        log(f"Resposta STRING publicada em '{reply_to}': {body}")
    except Exception:
        log(f"Erro ao publicar resposta STRING: {traceback.format_exc()}")


def publish_read_file_response(channel, reply_to: str, response: dict,
                               correlation_id: str = "") -> None:
    """Publica resposta de query.read.file.v1: {fileName, content} ou {error}."""
    try:
        body = json.dumps(response, ensure_ascii=False)
        props = pika.BasicProperties(
            delivery_mode=2,
            correlation_id=correlation_id if correlation_id else None,
            content_type="application/json",
        )
        channel.basic_publish(
            exchange="", routing_key=reply_to,
            body=body.encode("utf-8"),
            properties=props,
        )
        log(f"Resposta READ FILE publicada em '{reply_to}': {body[:200]}")
    except Exception:
        log(f"Erro ao publicar resposta READ FILE: {traceback.format_exc()}")


# ------------------------------------------------------------------ #
#  Montagem de argumentos
# ------------------------------------------------------------------ #
def build_args_v1(routing_key: str, payload: dict) -> list[str] | None:
    """
    Converte payload JSON nos args posicionais esperados pelo VBS.
    Retorna None se o routing_key nao tiver mapeamento de args.
    """
    file_name = payload.get("fileName", "").strip().upper()
    content   = payload.get("fileContent", payload.get("content", ""))
    encoded   = content.replace("\r\n", "\\n").replace("\r", "\\n").replace("\n", "\\n")

    if routing_key in ("command.fix.v1", "command.revert.v1", "command.create.v1"):
        return [file_name, encoded]

    if routing_key == "command.buscar.v1":
        return [file_name]

    if routing_key == "query.requests.v1":
        return []

    if routing_key == "query.all.files.v1":
        return []

    if routing_key == "query.all.packages.v1":
        return []

    if routing_key == "query.versions.metadata.v1":
        file_name = payload.get("fileName", "").strip().upper()
        category  = payload.get("category", "").strip().upper()
        return [file_name, category]

    if routing_key == "query.file.category.v1":
        file_name = payload.get("fileName", "").strip().upper()
        return [file_name]

    if routing_key == "query.request.files.v1":
        requests = payload.get("requests", [])
        if isinstance(requests, list):
            return [",".join(r.strip() for r in requests if r.strip())]
        return [str(requests).strip()]

    if routing_key == "query.request.description.v1":
        request_id = payload.get("requestId", "").strip().upper()
        return [request_id]

    if routing_key == "query.read.from.version.v1":
        file_name  = payload.get("fileName", "").strip().upper()
        category   = payload.get("category", "").strip().upper()
        version_id = payload.get("versionId", "").strip()
        return [file_name, category, version_id]

    return None


# ------------------------------------------------------------------ #
#  Helper genérico — download + execução + tratamento de erro
# ------------------------------------------------------------------ #
def _run_query_vbs(channel, payload: dict, vbs_url: str,
                   args: list[str], routing_key: str):
    """
    Executa download + cscript + tratamento padrao de erro.

    Retorna (True, stdout, correlation_id, reply_to)  em caso de sucesso
    ou      (False, None, correlation_id, reply_to) em caso de falha
    (nesse caso ja publica a mensagem de erro no reply_to).
    """
    correlation_id = payload.get("correlationId", "")
    reply_to       = payload.get("replyTo", QUEUE_RESPONSES)

    log(f"[QUERY] {routing_key} | args={args} | correlationId={correlation_id}")

    vbs_path = download_vbs(vbs_url)
    if vbs_path is None:
        error_msg = "ERRO: Falha no download do VBS"
        log(f"[QUERY] FALHA: {routing_key}: {error_msg}")
        publish_string_response(channel, reply_to, error_msg, correlation_id)
        return False, None, correlation_id, reply_to

    ok, details = execute_vbs(vbs_path, args)

    if not ok:
        error_msg = f"ERRO: {details}" if details else "ERRO: Falha ao executar VBS"
        log(f"[QUERY] FALHA: {routing_key}: {details}")
        publish_string_response(channel, reply_to, error_msg, correlation_id)
        return False, None, correlation_id, reply_to

    return True, details, correlation_id, reply_to


# ------------------------------------------------------------------ #
#  Handlers de query
# ------------------------------------------------------------------ #
def process_query_requests(channel, payload: dict, vbs_url: str) -> None:
    ok, details, cid, reply = _run_query_vbs(channel, payload, vbs_url, [], "query.requests.v1")
    if not ok:
        return
    try:
        data = parse_requests_txt(details)
    except Exception as exc:
        log(f"[QUERY] Erro no parsing do TXT: {exc}")
        publish_string_response(channel, reply, f"ERRO: Falha no parsing: {exc}", cid)
        return
    log(f"[QUERY] SUCESSO: {len(data)} requests encontradas")
    publish_query_response(channel, reply, data, cid)


def process_query_all_files(channel, payload: dict, vbs_url: str) -> None:
    ok, details, cid, reply = _run_query_vbs(channel, payload, vbs_url, [], "query.all.files.v1")
    if not ok:
        return
    try:
        data = parse_reports_txt(details)
    except Exception as exc:
        log(f"[QUERY] Erro no parsing do TXT: {exc}")
        publish_string_response(channel, reply, f"ERRO: Falha no parsing: {exc}", cid)
        return
    log(f"[QUERY] SUCESSO: {len(data)} reports encontrados")
    publish_query_response(channel, reply, data, cid)


def process_query_all_packages(channel, payload: dict, vbs_url: str) -> None:
    ok, details, cid, reply = _run_query_vbs(channel, payload, vbs_url, [], "query.all.packages.v1")
    if not ok:
        return
    try:
        data = parse_packages_txt(details)
    except Exception as exc:
        log(f"[QUERY] Erro no parsing do TXT: {exc}")
        publish_string_response(channel, reply, f"ERRO: Falha no parsing: {exc}", cid)
        return
    log(f"[QUERY] SUCESSO: {len(data)} pacotes encontrados")
    publish_query_response(channel, reply, data, cid)


def process_query_versions_metadata(channel, payload: dict, vbs_url: str) -> None:
    file_name = payload.get("fileName", "").strip().upper()
    category  = payload.get("category", "").strip().upper()
    if not file_name:
        cid = payload.get("correlationId", "")
        reply = payload.get("replyTo", QUEUE_RESPONSES)
        publish_string_response(channel, reply, "ERRO: fileName ausente no payload", cid)
        return
    ok, details, cid, reply = _run_query_vbs(
        channel, payload, vbs_url, [file_name, category], "query.versions.metadata.v1")
    if not ok:
        return
    try:
        data = parse_versions_metadata_txt(details)
    except Exception as exc:
        log(f"[QUERY] Erro no parsing do TXT: {exc}")
        publish_string_response(channel, reply, f"ERRO: Falha no parsing: {exc}", cid)
        return
    log(f"[QUERY] SUCESSO: {len(data)} versoes encontradas")
    publish_query_response(channel, reply, data, cid)


def process_query_file_category(channel, payload: dict, vbs_url: str) -> None:
    file_name = payload.get("fileName", "").strip().upper()
    if not file_name:
        cid = payload.get("correlationId", "")
        reply = payload.get("replyTo", QUEUE_RESPONSES)
        publish_string_response(channel, reply, "ERRO: fileName ausente no payload", cid)
        return
    ok, details, cid, reply = _run_query_vbs(
        channel, payload, vbs_url, [file_name], "query.file.category.v1")
    if not ok:
        return
    category = details.replace("\\n", "").strip()
    log(f"[QUERY] SUCESSO: query.file.category.v1 - {file_name} = {category}")
    publish_string_response(channel, reply, category, cid)


def process_query_request_files(channel, payload: dict, vbs_url: str) -> None:
    requests = payload.get("requests", [])
    if isinstance(requests, list):
        requests_str = ",".join(r.strip() for r in requests if r.strip())
    else:
        requests_str = str(requests).strip()
    if not requests_str:
        cid = payload.get("correlationId", "")
        reply = payload.get("replyTo", QUEUE_RESPONSES)
        publish_string_response(channel, reply, "ERRO: requests ausente no payload", cid)
        return
    ok, details, cid, reply = _run_query_vbs(
        channel, payload, vbs_url, [requests_str], "query.request.files.v1")
    if not ok:
        return
    try:
        data = parse_abap_files_by_request_txt(details)
    except Exception as exc:
        log(f"[QUERY] Erro no parsing do TXT: {exc}")
        publish_string_response(channel, reply, f"ERRO: Falha no parsing: {exc}", cid)
        return
    log(f"[QUERY] SUCESSO: {len(data)} arquivos encontrados para requests")
    publish_query_response(channel, reply, data, cid)


def process_query_request_description(channel, payload: dict, vbs_url: str) -> None:
    request_id = payload.get("requestId", "").strip().upper()
    if not request_id:
        cid = payload.get("correlationId", "")
        reply = payload.get("replyTo", QUEUE_RESPONSES)
        publish_string_response(channel, reply, "ERRO: requestId ausente no payload", cid)
        return
    ok, details, cid, reply = _run_query_vbs(
        channel, payload, vbs_url, [request_id], "query.request.description.v1")
    if not ok:
        return
    description = details.replace("\\n", "").strip()
    if description.startswith('"') and description.endswith('"'):
        description = description[1:-1]
    log(f"[QUERY] SUCESSO: query.request.description.v1 - {request_id} = {description}")
    publish_string_response(channel, reply, description, cid)


def process_query_read_from_version(channel, payload: dict, vbs_url: str) -> None:
    file_name  = payload.get("fileName", "").strip().upper()
    category   = payload.get("category", "").strip().upper()
    version_id = payload.get("versionId", "").strip()
    if not file_name:
        cid = payload.get("correlationId", "")
        reply = payload.get("replyTo", QUEUE_RESPONSES)
        publish_string_response(channel, reply, "ERRO: fileName ausente no payload", cid)
        return
    if not version_id:
        cid = payload.get("correlationId", "")
        reply = payload.get("replyTo", QUEUE_RESPONSES)
        publish_string_response(channel, reply, "ERRO: versionId ausente no payload", cid)
        return
    ok, details, cid, reply = _run_query_vbs(
        channel, payload, vbs_url, [file_name, category, version_id], "query.read.from.version.v1")
    if not ok:
        return
    content = details.replace("\\n", "\n")
    log(f"[QUERY] SUCESSO: query.read.from.version.v1 - {file_name} v{version_id} ({len(content)} chars)")
    publish_string_response(channel, reply, content, cid)


def process_query_read_file(channel, payload: dict, vbs_url: str) -> None:
    correlation_id = payload.get("correlationId", "")
    reply_to       = payload.get("replyTo", QUEUE_RESPONSES)
    file_name      = payload.get("fileName", "").strip().upper()

    log(f"[QUERY] query.read.file.v1 | fileName={file_name} | correlationId={correlation_id}")

    if not file_name:
        publish_read_file_response(channel, reply_to, {"error": "fileName ausente no payload"}, correlation_id)
        return

    vbs_path = download_vbs(vbs_url)
    if vbs_path is None:
        publish_read_file_response(channel, reply_to, {"error": "Falha no download do VBS"}, correlation_id)
        return

    ok, details = execute_vbs(vbs_path, [file_name])

    if ok:
        log(f"[QUERY] SUCESSO: query.read.file.v1 - {file_name}")
        publish_read_file_response(
            channel, reply_to, {"fileName": file_name, "content": details}, correlation_id
        )
    else:
        log(f"[QUERY] FALHA: query.read.file.v1 - {file_name}: {details}")
        publish_read_file_response(
            channel, reply_to, {"error": details or "Erro ao executar buscarProgSap"}, correlation_id
        )


# ------------------------------------------------------------------ #
#  Handler generico de comando
# ------------------------------------------------------------------ #
def process_v1_event(channel, payload: dict, routing_key: str, vbs_url: str) -> None:
    correlation_id = payload.get("correlationId", "")
    reply_to       = payload.get("replyTo", QUEUE_RESPONSES)
    file_name      = payload.get("fileName", "desconhecido").strip().upper()

    log(f"[V1] routing_key={routing_key} | fileName={file_name} | correlationId={correlation_id}")

    args = build_args_v1(routing_key, payload)
    if args is None:
        log(f"[V1] Sem mapeamento de args para routing_key '{routing_key}'.")
        publish_fix_response(channel, reply_to, True,
                             f"routing_key '{routing_key}' sem mapeamento de argumentos.",
                             False, correlation_id)
        return

    if not args[0]:
        publish_fix_response(channel, reply_to, True,
                             "fileName ausente no payload.", False, correlation_id)
        return

    vbs_path = download_vbs(vbs_url)
    if vbs_path is None:
        publish_fix_response(channel, reply_to, True,
                             f"Falha no download do VBS: {vbs_url}", False, correlation_id)
        return

    ok, details = execute_vbs(vbs_path, args)

    if ok:
        log(f"[V1] SUCESSO: {routing_key} - {file_name}")
        publish_fix_response(channel, reply_to, False, "", False, correlation_id)
    else:
        log(f"[V1] FALHA: {routing_key} - {file_name}: {details}")
        publish_fix_response(channel, reply_to, True,
                             details or "Erro ao executar script SAP GUI.",
                             False, correlation_id)


# ------------------------------------------------------------------ #
#  Callback principal v1
# ------------------------------------------------------------------ #
def callback_v1(ch, method, properties, body):
    """Callback das filas v1 — roteia pelo header amqp_receivedRoutingKey."""
    headers     = getattr(properties, "headers", None) or {}
    received_rk = headers.get("amqp_receivedRoutingKey", "") or method.routing_key or ""
    body_str    = body.decode("utf-8")

    command = received_rk.removeprefix(ROUTING_KEY_PREFIX) if received_rk else ""
    log(f"=== EVENTO V1 === received_rk={received_rk} | command={command} | {body_str}")

    try:
        payload = json.loads(body_str)
    except json.JSONDecodeError as e:
        log(f"JSON invalido no evento V1: {e}")
        return

    if not isinstance(payload, dict):
        payload = {}

    # Ignora mensagens que sao respostas (evita self-consumption loop)
    if "hasError" in payload and "isCodeError" in payload:
        log(f"[V1] Mensagem e uma resposta de fix (nao comando). Ignorada.")
        return

    reply_to = (
        getattr(properties, "reply_to", None)
        or payload.get("replyTo")
        or QUEUE_RESPONSES
    )
    correlation_id = (
        getattr(properties, "correlation_id", None)
        or payload.get("correlationId", "")
    )
    payload["replyTo"]       = reply_to
    payload["correlationId"] = correlation_id

    log(f"[V1] replyTo={reply_to} | correlationId={correlation_id}")

    if not command:
        log(f"[V1] Header 'amqp_receivedRoutingKey' ausente ou vazio. Mensagem ignorada.")
        return

    vbs_url = VBS_BY_ROUTING_KEY.get(command)
    if vbs_url is None:
        log(f"[V1] Comando '{command}' nao mapeado em VBS_BY_ROUTING_KEY. Ignorado.")
        return

    try:
        if command == "query.read.file.v1":
            process_query_read_file(ch, payload, vbs_url)
        elif command == "query.requests.v1":
            process_query_requests(ch, payload, vbs_url)
        elif command == "query.all.files.v1":
            process_query_all_files(ch, payload, vbs_url)
        elif command == "query.all.packages.v1":
            process_query_all_packages(ch, payload, vbs_url)
        elif command == "query.versions.metadata.v1":
            process_query_versions_metadata(ch, payload, vbs_url)
        elif command == "query.file.category.v1":
            process_query_file_category(ch, payload, vbs_url)
        elif command == "query.request.files.v1":
            process_query_request_files(ch, payload, vbs_url)
        elif command == "query.request.description.v1":
            process_query_request_description(ch, payload, vbs_url)
        elif command == "query.read.from.version.v1":
            process_query_read_from_version(ch, payload, vbs_url)
        elif command in QUERY_ROUTING_KEYS:
            log(f"[V1] Query '{command}' sem handler dedicado. Ignorado.")
        else:
            process_v1_event(ch, payload, command, vbs_url)
    except Exception:
        log(f"Erro nao tratado (v1): {traceback.format_exc()}")
