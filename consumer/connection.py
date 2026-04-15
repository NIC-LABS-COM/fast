"""Conexao RabbitMQ."""

import ssl
import pika

from .config import BROKER_HOST, BROKER_PORT, BROKER_USER, BROKER_PASSWORD, BROKER_VHOST, BROKER_TIMEOUT


def get_rabbitmq_connection():
    credentials = pika.PlainCredentials(BROKER_USER, BROKER_PASSWORD)

    ssl_context = ssl.create_default_context()
    ssl_options = pika.SSLOptions(ssl_context, BROKER_HOST)

    params = pika.ConnectionParameters(
        host=BROKER_HOST, port=BROKER_PORT, virtual_host=BROKER_VHOST,
        credentials=credentials,
        ssl_options=ssl_options,
        blocked_connection_timeout=BROKER_TIMEOUT / 1000,
    )
    return pika.BlockingConnection(params)
