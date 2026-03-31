"""Conexao RabbitMQ."""

import ssl

import pika

from .config import RABBITMQ_HOST, RABBITMQ_USER, RABBITMQ_PASS, RABBITMQ_VHOST


def get_rabbitmq_connection():
    credentials = pika.PlainCredentials(RABBITMQ_USER, RABBITMQ_PASS)
    context = ssl.create_default_context()
    params = pika.ConnectionParameters(
        host=RABBITMQ_HOST, port=5671, virtual_host=RABBITMQ_VHOST,
        credentials=credentials, ssl_options=pika.SSLOptions(context),
    )
    return pika.BlockingConnection(params)
