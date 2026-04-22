"""Conexao RabbitMQ com retry e exponential backoff."""

import ssl
import socket
import time
import pika

from .config import (
    BROKERS,
    MAX_RECONNECT_ATTEMPTS, RECONNECT_BACKOFF_SECONDS,
)
from .logger import log


def _diagnose_error(exc: Exception, broker: dict) -> str:
    """Retorna mensagem descritiva do erro de conexao."""
    name = broker.get("name", broker["host"])
    host = broker["host"]
    port = broker["port"]
    user = broker["user"]
    msg = str(exc)

    # Erro de autenticacao (403)
    if isinstance(exc, pika.exceptions.ProbableAuthenticationError) or "ACCESS_REFUSED" in msg or "403" in msg:
        return (f"[DIAGNOSTICO][{name}] FALHA DE AUTENTICACAO no broker {host}:{port}.\n"
                f"  -> Usuario '{user}' foi recusado pelo broker.\n"
                f"  -> Verifique: 1) usuario/senha estao corretos  "
                f"2) usuario tem permissao no vhost '{broker['vhost']}'  "
                f"3) usuario esta habilitado no RabbitMQ Management.")

    # Erro de conexao recusada
    if isinstance(exc, pika.exceptions.AMQPConnectionError) or "Connection refused" in msg:
        return (f"[DIAGNOSTICO][{name}] CONEXAO RECUSADA em {host}:{port}.\n"
                f"  -> O broker nao esta aceitando conexoes nesta porta.\n"
                f"  -> Verifique: 1) o broker esta rodando  "
                f"2) a porta {port} esta correta  "
                f"3) firewall/VPN permite acesso ao host.")

    # Erro de DNS / host nao encontrado
    if isinstance(exc, socket.gaierror) or "Name or service not known" in msg or "getaddrinfo" in msg:
        return (f"[DIAGNOSTICO][{name}] HOST NAO ENCONTRADO: {host}.\n"
                f"  -> DNS nao conseguiu resolver o nome do host.\n"
                f"  -> Verifique: 1) o hostname esta correto  "
                f"2) voce tem acesso a rede/VPN  "
                f"3) DNS esta funcionando.")

    # Timeout
    if isinstance(exc, (socket.timeout, pika.exceptions.AMQPConnectionError)) and "timeout" in msg.lower():
        return (f"[DIAGNOSTICO][{name}] TIMEOUT ao conectar em {host}:{port}.\n"
                f"  -> O broker nao respondeu a tempo.\n"
                f"  -> Verifique: 1) rede/VPN esta estavel  "
                f"2) o broker esta sobrecarregado  "
                f"3) firewall nao esta bloqueando.")

    # Erro SSL
    if isinstance(exc, ssl.SSLError) or "SSL" in msg or "CERTIFICATE" in msg:
        return (f"[DIAGNOSTICO][{name}] ERRO SSL ao conectar em {host}:{port}.\n"
                f"  -> Falha no handshake TLS/SSL.\n"
                f"  -> Verifique: 1) a porta {port} usa SSL  "
                f"2) certificado do broker e valido  "
                f"3) versao TLS e compativel.")

    # Erro generico
    return (f"[DIAGNOSTICO][{name}] ERRO DESCONHECIDO ao conectar em {host}:{port}.\n"
            f"  -> Tipo: {type(exc).__name__}\n"
            f"  -> Detalhes: {msg}")


def _build_connection_params(broker: dict) -> pika.ConnectionParameters:
    """Constrói os parâmetros de conexão RabbitMQ para um broker."""
    credentials = pika.PlainCredentials(broker["user"], broker["password"])
    ssl_context = ssl.create_default_context()
    ssl_options = pika.SSLOptions(ssl_context, broker["host"])
    return pika.ConnectionParameters(
        host=broker["host"],
        port=broker["port"],
        virtual_host=broker["vhost"],
        credentials=credentials,
        ssl_options=ssl_options,
        blocked_connection_timeout=broker["timeout"] / 1000,
    )


def get_rabbitmq_connection(broker: dict, max_attempts=None, backoff=None):
    """Cria conexão RabbitMQ com retry e exponential backoff.

    Tenta conectar até max_attempts vezes com backoff exponencial
    entre tentativas. Retorna pika.BlockingConnection ou levanta
    a última exceção se todas as tentativas falharem.
    """
    if max_attempts is None:
        max_attempts = MAX_RECONNECT_ATTEMPTS
    if backoff is None:
        backoff = RECONNECT_BACKOFF_SECONDS

    name = broker.get("name", broker["host"])
    params = _build_connection_params(broker)
    last_error = None

    for attempt in range(1, max_attempts + 1):
        try:
            log(f"[CONN][{name}] Tentativa {attempt}/{max_attempts} de conexao "
                f"com {broker['host']}:{broker['port']}...")
            connection = pika.BlockingConnection(params)
            log(f"[CONN][{name}] Conectado com sucesso ao RabbitMQ.")
            return connection
        except Exception as exc:
            last_error = exc
            log(f"[CONN][{name}] Tentativa {attempt}/{max_attempts} falhou: {exc}")
            log(_diagnose_error(exc, broker))
            if attempt < max_attempts:
                wait = min(backoff * (2 ** (attempt - 1)), 60)
                log(f"[CONN][{name}] Aguardando {wait:.0f}s antes da proxima "
                    f"tentativa...")
                time.sleep(wait)

    log(f"[CONN][{name}] TODAS as {max_attempts} tentativas de conexao falharam.")
    log(f"[CONN][{name}] Ultimo erro: {type(last_error).__name__}: {last_error}")
    raise last_error
