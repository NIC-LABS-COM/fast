"""
Servico de comunicacao com RabbitMQ.
Responsavel por publicar comandos e escutar respostas.
"""
import json
import logging
import ssl
import threading
import time
import traceback
from typing import Callable, Dict, List, Tuple

import pika

from core.config import (
    QUEUE_COMMANDS,
    QUEUE_RESPONSES,
    RABBITMQ_HOST,
    RABBITMQ_PASS,
    RABBITMQ_USER,
    RABBITMQ_VHOST,
)

logger = logging.getLogger("sap_publisher")


class RabbitMQService:
    """Encapsula todas as operacoes com o broker RabbitMQ."""

    def _get_connection(self) -> pika.BlockingConnection:
        credentials = pika.PlainCredentials(RABBITMQ_USER, RABBITMQ_PASS)
        context = ssl.create_default_context()
        params = pika.ConnectionParameters(
            host=RABBITMQ_HOST,
            port=5671,
            virtual_host=RABBITMQ_VHOST,
            credentials=credentials,
            ssl_options=pika.SSLOptions(context),
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
