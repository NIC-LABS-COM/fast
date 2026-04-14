"""Conexao RabbitMQ."""

import pika

from .config import BROKER_HOST, BROKER_PORT, BROKER_USER, BROKER_PASSWORD, BROKER_VHOST, BROKER_TIMEOUT


def get_rabbitmq_connection():
    credentials = pika.PlainCredentials(BROKER_USER, BROKER_PASSWORD)
    params = pika.ConnectionParameters(
        host=BROKER_HOST, port=BROKER_PORT, virtual_host=BROKER_VHOST,
        credentials=credentials,
        blocked_connection_timeout=BROKER_TIMEOUT / 1000,
    )
    return pika.BlockingConnection(params)
