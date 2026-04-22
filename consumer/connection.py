"""Conexao RabbitMQ com retry e exponential backoff."""

import ssl
import time
import pika

from .config import (
    BROKER_HOST, BROKER_PORT, BROKER_USER, BROKER_PASSWORD,
    BROKER_VHOST, BROKER_TIMEOUT,
    MAX_RECONNECT_ATTEMPTS, RECONNECT_BACKOFF_SECONDS,
)
from .logger import log


def _build_connection_params() -> pika.ConnectionParameters:
    """Constrói os parâmetros de conexão RabbitMQ."""
    credentials = pika.PlainCredentials(BROKER_USER, BROKER_PASSWORD)
    ssl_context = ssl.create_default_context()
    ssl_options = pika.SSLOptions(ssl_context, BROKER_HOST)
    return pika.ConnectionParameters(
        host=BROKER_HOST,
        port=BROKER_PORT,
        virtual_host=BROKER_VHOST,
        credentials=credentials,
        ssl_options=ssl_options,
        blocked_connection_timeout=BROKER_TIMEOUT / 1000,
    )


def get_rabbitmq_connection(max_attempts=None, backoff=None):
    """Cria conexão RabbitMQ com retry e exponential backoff.

    Tenta conectar até max_attempts vezes com backoff exponencial
    entre tentativas. Retorna pika.BlockingConnection ou levanta
    a última exceção se todas as tentativas falharem.
    """
    if max_attempts is None:
        max_attempts = MAX_RECONNECT_ATTEMPTS
    if backoff is None:
        backoff = RECONNECT_BACKOFF_SECONDS

    params = _build_connection_params()
    last_error = None

    for attempt in range(1, max_attempts + 1):
        try:
            log(f"[CONN] Tentativa {attempt}/{max_attempts} de conexao "
                f"com {BROKER_HOST}:{BROKER_PORT}...")
            connection = pika.BlockingConnection(params)
            log(f"[CONN] Conectado com sucesso ao RabbitMQ.")
            return connection
        except Exception as exc:
            last_error = exc
            log(f"[CONN] Tentativa {attempt}/{max_attempts} falhou: {exc}")
            if attempt < max_attempts:
                wait = min(backoff * (2 ** (attempt - 1)), 60)
                log(f"[CONN] Aguardando {wait:.0f}s antes da proxima "
                    f"tentativa...")
                time.sleep(wait)

    log(f"[CONN] TODAS as {max_attempts} tentativas de conexao falharam.")
    raise last_error
