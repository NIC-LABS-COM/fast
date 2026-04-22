"""
ConsumerTT — Consumer RabbitMQ para automacao SAP GUI.

Conecta simultaneamente a multiplos brokers (PROD, DEV, HOM) via threads.
Cada broker tem sua propria conexão, canal e loop de reconexão.
As respostas voltam pelo mesmo canal/broker (via replyTo).

Escuta duas filas em paralelo por broker:
  1. queue_vpn_usiminas  — arquitetura legada (payload com action/vbs_url/args)
  2. q.usiminas.v1       — arquitetura orientada a eventos (routing key + payload tipado)

Resiliência:
  - Dead Letter Exchange (DLX) + Dead Letter Queue (DLQ) para mensagens com falha
  - Retry com limite de tentativas e backoff (equivalente ao Spring RetryInterceptor)
  - Manual ACK/NACK para garantir entrega confiável (sem auto_ack)
  - Reconexão automática com backoff em caso de queda de conexão
"""

import os
import time
import threading
import traceback

import pika

from .config import (
    QUEUE_COMMANDS, QUEUE_RESPONSES, QUEUES_V1,
    EXCHANGE_V1, ROUTING_KEY_BIND, BROKERS,
    VBS_DIR, VBS_BY_ROUTING_KEY,
    DEAD_LETTER_EXCHANGE, QUEUE_V1_DLQ,
    MAX_RETRY_ATTEMPTS, RETRY_BACKOFF_SECONDS,
)
from .connection import get_rabbitmq_connection
from .logger import log
from .legacy import callback_legacy
from .v1 import callback_v1
from .retry import with_retry_and_ack


# Intervalo (segundos) antes de tentar reconectar após queda de conexão
_RECONNECT_DELAY = 10


def _check_vbs_files() -> None:
    log(f"VBS_DIR: {VBS_DIR}")
    mapped = set(VBS_BY_ROUTING_KEY.values())

    # Arquivos presentes fisicamente no diretório
    try:
        found = {f for f in os.listdir(VBS_DIR) if f.lower().endswith(".vbs")}
    except FileNotFoundError:
        log(f"  [ERRO] Diretório VBS não encontrado: {VBS_DIR}")
        return

    # Mapeados e presentes
    ok = sorted(mapped & found)
    # Mapeados mas ausentes
    missing = sorted(mapped - found)
    # Presentes mas sem routing key (apenas frontend ou extra)
    extra = sorted(found - mapped)

    for name in ok:
        log(f"  [OK]       {name}")
    for name in missing:
        log(f"  [FALTANDO] {name}")
    for name in extra:
        log(f"  [EXTRA]    {name}  (sem routing key no consumer)")

    log(f"VBS: {len(ok)} mapeados OK, {len(missing)} faltando, {len(extra)} extras.")


def _declare_dead_letter_infrastructure(channel) -> None:
    """Declara Dead Letter Exchange (DLX) e Dead Letter Queues (DLQ).

    Equivalente aos beans deadLetterExchange() + sapDlq() +
    deadLetterBinding() do RabbitMQConfig Spring.
    """
    # DLX — DirectExchange, durable (equivalente ao Spring DirectExchange)
    channel.exchange_declare(
        exchange=DEAD_LETTER_EXCHANGE,
        exchange_type="direct",
        durable=True,
        auto_delete=False,
    )
    log(f"DLX declarado: {DEAD_LETTER_EXCHANGE} (direct, durable)")

    # DLQ para cada fila V1
    for queue in QUEUES_V1:
        dlq_name = f"{queue}.dlq"
        channel.queue_declare(queue=dlq_name, durable=True)
        channel.queue_bind(
            queue=dlq_name,
            exchange=DEAD_LETTER_EXCHANGE,
            routing_key=dlq_name,
        )
        log(f"DLQ declarada e vinculada: {dlq_name} -> {DEAD_LETTER_EXCHANGE}")


def _declare_v1_queue(connection, channel, queue: str):
    """Declara fila V1 com argumentos de Dead Letter.

    Equivalente ao QueueBuilder.durable(...)
        .deadLetterExchange(DLX)
        .deadLetterRoutingKey(dlqRoutingKey)
        .build() do Spring.

    Se a fila já existir no broker SEM os argumentos de DLQ
    (situação de migração), faz fallback para declaração simples
    e loga aviso para o operador.
    """
    dlq_name = f"{queue}.dlq"
    dlq_args = {
        "x-dead-letter-exchange": DEAD_LETTER_EXCHANGE,
        "x-dead-letter-routing-key": dlq_name,
    }

    try:
        channel.queue_declare(queue=queue, durable=True, arguments=dlq_args)
        log(f"Fila declarada com DLQ: {queue} "
            f"(x-dead-letter-exchange={DEAD_LETTER_EXCHANGE})")
        return channel
    except pika.exceptions.ChannelClosedByBroker as exc:
        if exc.reply_code == 406:
            log(f"[AVISO] Fila '{queue}' ja existe com argumentos diferentes.")
            log(f"[AVISO] Para habilitar DLQ, delete a fila no broker e "
                f"reinicie o consumer.")
            log(f"[AVISO] Declarando sem argumentos DLQ (compatibilidade)...")
            # Canal foi fechado pelo broker — cria um novo
            channel = connection.channel()
            channel.queue_declare(queue=queue, durable=True)
            log(f"Fila declarada (sem DLQ): {queue}")
            return channel
        raise


def _setup_channel(connection, broker_name: str):
    """Configura canal com toda a infraestrutura de filas, exchanges e bindings."""
    channel = connection.channel()

    # 1. Infraestrutura DLQ (deve vir antes das filas principais)
    _declare_dead_letter_infrastructure(channel)

    # 2. Filas legado (sem DLQ — mantém compatibilidade)
    channel.queue_declare(queue=QUEUE_COMMANDS, durable=True)
    channel.queue_declare(queue=QUEUE_RESPONSES, durable=True)

    # 3. Exchange V1 (topic, durable)
    channel.exchange_declare(
        exchange=EXCHANGE_V1, exchange_type="topic",
        durable=True, auto_delete=False,
    )
    log(f"[{broker_name}] Exchange declarado/verificado: {EXCHANGE_V1} "
        f"(topic, durable, auto_delete=False)")

    # 4. Filas V1 com DLQ args + binding
    for queue in QUEUES_V1:
        channel = _declare_v1_queue(connection, channel, queue)
        channel.queue_bind(
            queue=queue,
            exchange=EXCHANGE_V1,
            routing_key=ROUTING_KEY_BIND,
        )
        log(f"[{broker_name}] Binding registrado: {EXCHANGE_V1} -> {queue} "
            f"[{ROUTING_KEY_BIND}]")

    # 5. QoS — processa uma mensagem por vez
    channel.basic_qos(prefetch_count=1)

    # 6. Callbacks com retry + manual ack/nack
    safe_callback_v1 = with_retry_and_ack()(callback_v1)
    safe_callback_legacy = with_retry_and_ack()(callback_legacy)

    for queue in QUEUES_V1:
        channel.basic_consume(
            queue=queue,
            on_message_callback=safe_callback_v1,
            auto_ack=False,
        )

    channel.basic_consume(
        queue=QUEUE_COMMANDS,
        on_message_callback=safe_callback_legacy,
        auto_ack=False,
    )

    log(f"[{broker_name}] Escutando filas: '{QUEUE_COMMANDS}' e "
        f"'{', '.join(QUEUES_V1)}' (auto_ack=False, "
        f"max_retries={MAX_RETRY_ATTEMPTS})")
    return channel


def _broker_loop(broker: dict) -> None:
    """Loop de reconexão para um único broker. Roda em thread própria."""
    name = broker.get("name", broker["host"])

    while True:
        connection = None
        try:
            connection = get_rabbitmq_connection(broker)
            channel = _setup_channel(connection, name)
            log(f"[{name}] Aguardando mensagens...")
            channel.start_consuming()

        except KeyboardInterrupt:
            log(f"[{name}] Consumer encerrado pelo usuario.")
            try:
                channel.stop_consuming()
            except Exception:
                pass
            break

        except Exception as exc:
            log(f"[CONN][{name}] Erro no consumo: {type(exc).__name__}: {exc}")
            log(f"[CONN][{name}] Stacktrace:\n{traceback.format_exc()}")
            log(f"[CONN][{name}] A thread vai tentar reconectar em {_RECONNECT_DELAY}s... "
                f"(as outras threads nao sao afetadas)")
            time.sleep(_RECONNECT_DELAY)

        finally:
            if connection and not connection.is_closed:
                try:
                    connection.close()
                except Exception:
                    pass
                log(f"[{name}] Conexao encerrada.")


def main() -> None:
    log("##################################################")
    log("CONSUMER INICIADO — MULTI-BROKER")
    log(f"Brokers       : {', '.join(b['name'] for b in BROKERS)}")
    log(f"Fila legado   : {QUEUE_COMMANDS}")
    log(f"Filas v1      : {', '.join(QUEUES_V1)}")
    log(f"Exchange v1   : {EXCHANGE_V1}")
    log(f"Binding v1    : {ROUTING_KEY_BIND}")
    log(f"Fila respostas: {QUEUE_RESPONSES}")
    log(f"DLX           : {DEAD_LETTER_EXCHANGE}")
    log(f"DLQ v1        : {QUEUE_V1_DLQ}")
    log(f"Max retries   : {MAX_RETRY_ATTEMPTS}")
    log(f"Retry backoff : {RETRY_BACKOFF_SECONDS}s (linear)")
    for b in BROKERS:
        log(f"  [{b['name']}] {b['host']}:{b['port']}")
    log("##################################################")
    _check_vbs_files()
    log("##################################################")

    # Uma thread por broker — cada uma com seu loop de reconexão
    threads = []
    for broker in BROKERS:
        t = threading.Thread(
            target=_broker_loop,
            args=(broker,),
            name=f"consumer-{broker['name']}",
            daemon=True,
        )
        t.start()
        threads.append(t)
        log(f"Thread iniciada para broker [{broker['name']}]")

    # Aguarda todas as threads (Ctrl+C encerra via daemon=True)
    try:
        for t in threads:
            t.join()
    except KeyboardInterrupt:
        log("Consumer encerrado pelo usuario (Ctrl+C).")

    log("Consumer finalizado.")


if __name__ == "__main__":
    main()
