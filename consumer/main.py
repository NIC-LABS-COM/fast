"""
ConsumerTT — Consumer RabbitMQ para automacao SAP GUI.

Escuta duas filas em paralelo:
  1. queue_vpn_usiminas  — arquitetura legada (payload com action/vbs_url/args)
  2. q.usiminas.v1       — arquitetura orientada a eventos (routing key + payload tipado)
"""

from .config import (
    QUEUE_COMMANDS, QUEUE_RESPONSES, QUEUES_V1,
    EXCHANGE_V1, ROUTING_KEY_BIND, RABBITMQ_HOST,
)
from .connection import get_rabbitmq_connection
from .logger import log
from .legacy import callback_legacy
from .v1 import callback_v1


def main() -> None:
    log("##################################################")
    log("CONSUMER INICIADO")
    log(f"Fila legado   : {QUEUE_COMMANDS}")
    log(f"Filas v1      : {', '.join(QUEUES_V1)}")
    log(f"Exchange v1   : {EXCHANGE_V1}")
    log(f"Binding v1    : {ROUTING_KEY_BIND}")
    log(f"Fila respostas: {QUEUE_RESPONSES}")
    log(f"Host          : {RABBITMQ_HOST}")
    log("##################################################")

    connection = get_rabbitmq_connection()
    channel    = connection.channel()

    # Filas legado
    channel.queue_declare(queue=QUEUE_COMMANDS,  durable=True)
    channel.queue_declare(queue=QUEUE_RESPONSES, durable=True)

    # Exchange v1
    channel.exchange_declare(
        exchange=EXCHANGE_V1, exchange_type="topic",
        durable=True, auto_delete=False,
    )
    log(f"Exchange declarado/verificado: {EXCHANGE_V1} (topic, durable, auto_delete=False)")

    # Filas v1
    for queue in QUEUES_V1:
        channel.queue_declare(queue=queue, durable=True)
        channel.queue_bind(queue=queue, exchange=EXCHANGE_V1, routing_key=ROUTING_KEY_BIND)
        log(f"Binding registrado: {EXCHANGE_V1} -> {queue} [{ROUTING_KEY_BIND}]")
        channel.basic_consume(queue=queue, on_message_callback=callback_v1, auto_ack=True)

    channel.basic_qos(prefetch_count=1)

    # Consumer legado
    channel.basic_consume(
        queue=QUEUE_COMMANDS,
        on_message_callback=callback_legacy,
        auto_ack=True,
    )

    log(f"Escutando filas: '{QUEUE_COMMANDS}' e '{', '.join(QUEUES_V1)}' (CTRL+C para sair)")
    print(f"Aguardando mensagens... (CTRL+C para sair)")

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
