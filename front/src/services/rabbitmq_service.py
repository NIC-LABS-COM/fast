"""
Servico de comunicacao com RabbitMQ.
Responsavel por publicar comandos e escutar respostas.
"""
import json
import logging
import threading
import time
import traceback
from typing import Callable, Dict, List, Tuple

import pika

from core.config import (
    QUEUE_COMMANDS,
    QUEUE_RESPONSES,
    BROKER_HOST,
    BROKER_PORT,
    BROKER_PASSWORD,
    BROKER_USER,
    BROKER_VHOST,
    BROKER_TIMEOUT,
)

logger = logging.getLogger("sap_publisher")


class RabbitMQService:
    """Encapsula todas as operacoes com o broker RabbitMQ."""

    def _get_connection(self) -> pika.BlockingConnection:
        credentials = pika.PlainCredentials(BROKER_USER, BROKER_PASSWORD)
        params = pika.ConnectionParameters(
            host=BROKER_HOST,
            port=BROKER_PORT,
            virtual_host=BROKER_VHOST,
            credentials=credentials,
            blocked_connection_timeout=BROKER_TIMEOUT / 1000,
        )
        return pika.BlockingConnection(params)

    def publish(self, message: Dict) -> bool:
        """Publica uma unica mensagem na fila de comandos."""
        try:
            conn = self._get_connection()
            ch = conn.channel()
            ch.queue_declare(queue=QUEUE_COMMANDS, durable=True)
            body = json.dumps(message, ensure_ascii=False)
            ch.basic_publish(
                exchange="",
                routing_key=QUEUE_COMMANDS,
                body=body.encode("utf-8"),
                properties=pika.BasicProperties(delivery_mode=2),
            )
            conn.close()
            logger.info(f"Publicado: {body}")
            return True
        except Exception:
            logger.error(f"Erro ao publicar: {traceback.format_exc()}")
            return False

    def publish_batch(self, messages: List[Dict]) -> Tuple[int, int]:
        """Publica multiplas mensagens em uma unica conexao. Retorna (ok, fail)."""
        ok_count = 0
        fail_count = 0
        try:
            conn = self._get_connection()
            ch = conn.channel()
            ch.queue_declare(queue=QUEUE_COMMANDS, durable=True)
            for msg in messages:
                try:
                    body = json.dumps(msg, ensure_ascii=False)
                    ch.basic_publish(
                        exchange="",
                        routing_key=QUEUE_COMMANDS,
                        body=body.encode("utf-8"),
                        properties=pika.BasicProperties(delivery_mode=2),
                    )
                    logger.info(f"Publicado: {body}")
                    ok_count += 1
                except Exception:
                    logger.error(f"Erro na mensagem: {traceback.format_exc()}")
                    fail_count += 1
            conn.close()
        except Exception:
            logger.error(f"Erro de conexao no batch: {traceback.format_exc()}")
            fail_count = len(messages) - ok_count
        return ok_count, fail_count

    def publish_chain(
        self,
        messages: List[Dict],
        timeout: float = 120.0,
        on_step: Callable[[int, int, Dict], None] | None = None,
    ) -> List[Dict]:
        """
        Publica mensagens sequencialmente, esperando a resposta de cada uma
        antes de enviar a proxima. Para imediatamente se houver erro real.

        Args:
            messages: lista de payloads (formato legado com correlationId).
            timeout: tempo maximo (s) de espera por cada resposta.
            on_step: callback opcional (step_index, total, response) a cada resposta.

        Returns:
            Lista de respostas recebidas (uma por mensagem processada).
        """
        responses: List[Dict] = []
        total = len(messages)

        try:
            conn = self._get_connection()
            ch = conn.channel()
            ch.queue_declare(queue=QUEUE_COMMANDS, durable=True)
            ch.queue_declare(queue=QUEUE_RESPONSES, durable=True)
        except Exception:
            logger.error(f"Erro de conexao no chain: {traceback.format_exc()}")
            return responses

        for idx, msg in enumerate(messages):
            correlation_id = msg.get("correlationId", "")
            action = msg.get("action", "?")
            obj_name = msg.get("args", ["?"])[0] if msg.get("args") else "?"

            # Publica a mensagem
            try:
                body = json.dumps(msg, ensure_ascii=False)
                ch.basic_publish(
                    exchange="",
                    routing_key=QUEUE_COMMANDS,
                    body=body.encode("utf-8"),
                    properties=pika.BasicProperties(delivery_mode=2),
                )
                logger.info(f"[CHAIN {idx+1}/{total}] Publicado: {action} | {obj_name}")
            except Exception:
                logger.error(f"[CHAIN] Erro ao publicar msg {idx+1}: {traceback.format_exc()}")
                responses.append({
                    "correlationId": correlation_id,
                    "action": action,
                    "object_name": obj_name,
                    "status": "erro",
                    "message": "Falha ao publicar mensagem.",
                })
                break

            # Espera resposta com correlationId correspondente
            response = self._wait_for_response(ch, correlation_id, timeout)
            if response is None:
                response = {
                    "correlationId": correlation_id,
                    "action": action,
                    "object_name": obj_name,
                    "status": "erro",
                    "message": f"Timeout ({timeout}s) esperando resposta do consumer.",
                }

            responses.append(response)
            if on_step:
                try:
                    on_step(idx, total, response)
                except Exception:
                    pass

            # Se erro real, aborta a cadeia
            status = response.get("status", "")
            if status == "erro":
                logger.warning(
                    f"[CHAIN] Erro em {action} | {obj_name}: {response.get('message', '')}. "
                    f"Abortando cadeia ({idx+1}/{total})."
                )
                break

        try:
            conn.close()
        except Exception:
            pass

        return responses

    def _wait_for_response(
        self, channel, correlation_id: str, timeout: float
    ) -> Dict | None:
        """Consome mensagens da fila de respostas ate encontrar a com o correlationId."""
        import time as _time

        deadline = _time.monotonic() + timeout

        while _time.monotonic() < deadline:
            remaining = max(0.1, deadline - _time.monotonic())
            method, properties, body = (None, None, None)
            try:
                method, properties, body = channel.basic_get(
                    queue=QUEUE_RESPONSES, auto_ack=True
                )
            except Exception:
                logger.error(f"Erro ao consumir resposta: {traceback.format_exc()}")
                return None

            if method is None:
                _time.sleep(0.5)
                continue

            body_str = body.decode("utf-8") if body else ""
            try:
                resp = json.loads(body_str)
            except Exception:
                continue

            if isinstance(resp, dict) and resp.get("correlationId") == correlation_id:
                return resp
            # Resposta de outro correlationId — descarta (pertence a outra operacao)

        return None

    def publish_v1(self, routing_key: str, payload: Dict) -> bool:
        """
        Publica mensagem usando arquitetura V1 com exchange topic.
        routing_key: ex: "usiminas.req.query.read.file.v1"
        payload: dict com campos {fileName, content, correlationId, replyTo}
        """
        try:
            conn = self._get_connection()
            ch = conn.channel()
            
            # Exchange topic já deve existir no broker
            exchange = "x.to-client.topic"
            
            body = json.dumps(payload, ensure_ascii=False)
            ch.basic_publish(
                exchange=exchange,
                routing_key=routing_key,
                body=body.encode("utf-8"),
                properties=pika.BasicProperties(delivery_mode=2, content_type="application/json"),
            )
            conn.close()
            logger.info(f"Publicado V1: {routing_key} → {body}")
            return True
        except Exception:
            logger.error(f"Erro ao publicar V1: {traceback.format_exc()}")
            return False

    def start_listener(self, on_message: Callable[[Dict], None]) -> threading.Thread:
        """Inicia thread daemon que consome a fila de respostas indefinidamente."""
        thread = threading.Thread(
            target=self._listen_loop, args=(on_message,), daemon=True
        )
        thread.start()
        logger.info("Listener de respostas iniciado")
        return thread

    def _listen_loop(self, on_message: Callable[[Dict], None]) -> None:
        while True:
            try:
                conn = self._get_connection()
                ch = conn.channel()
                
                # Escuta fila legado
                ch.queue_declare(queue=QUEUE_RESPONSES, durable=True)
                
                def callback(ch, method, properties, body):
                    body_str = body.decode("utf-8")
                    try:
                        response = json.loads(body_str)
                        on_message(response)
                    except Exception:
                        logger.error(f"Erro ao parsear resposta: {body_str}")

                ch.basic_consume(
                    queue=QUEUE_RESPONSES,
                    on_message_callback=callback,
                    auto_ack=True,
                )
                ch.start_consuming()
            except Exception:
                logger.error(
                    f"Listener desconectou, reconectando em 5s... {traceback.format_exc()}"
                )
                time.sleep(5)
