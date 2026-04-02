"""Constantes e configuracao do Consumer RabbitMQ."""

import os
from datetime import datetime

# ------------------------------------------------------------------ #
#  Configuracao RabbitMQ
# ------------------------------------------------------------------ #
RABBITMQ_HOST  = "jackal.rmq.cloudamqp.com"
RABBITMQ_USER  = "rhrstugr"
RABBITMQ_PASS  = "HC2wvtBtou_DUk9AA276209T4718K9cF"
RABBITMQ_VHOST = "rhrstugr"

# Filas — legado
QUEUE_COMMANDS  = "queue_vpn_usiminas"
QUEUE_RESPONSES = "queue_vpn_respostas"

# Filas — nova arquitetura
QUEUE_V1       = "q.usiminas.v1"

# Filas V1 monitoradas (declarar, fazer bind e consumir todas)
QUEUES_V1 = [QUEUE_V1]

# Exchange — nova arquitetura
EXCHANGE_V1      = "x.to-client.topic"
ROUTING_KEY_BIND = "usiminas.req.#.v1"

# Mapeamento: routing_key -> URL do VBS no GitHub
VBS_BY_ROUTING_KEY: dict[str, str] = {
    "command.fix.v1":      "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/ScriptEditarSE38.vbs",
    "command.revert.v1":   "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/ScriptEditarSE38.vbs",
    "command.create.v1":   "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/ScriptCriarSE38.vbs",
    "command.buscar.v1":   "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscarProgSap.vbs",
    "query.read.file.v1":  "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/baixarArquivoSAP.vbs",
    "query.requests.v1":      "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscarRequests.vbs",
    "query.all.files.v1":     "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscaReports.vbs",
    "query.all.packages.v1":          "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscaPacotes.vbs",
    "query.versions.metadata.v1":     "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscaVersionsMetadata.vbs",
    "query.file.category.v1":         "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscaCategoryByFileName.vbs",
    "query.request.files.v1":         "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscaAbapFilesByRequest.vbs",
    "query.request.description.v1":   "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscaRequestDescription.vbs",
    "query.read.from.version.v1":      "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/vbs/buscaConteudoPorVersao.vbs",
}

# Routing keys que sao queries (retornam dados estruturados)
QUERY_ROUTING_KEYS: set[str] = {
    "query.requests.v1",
    "query.all.files.v1",
    "query.all.packages.v1",
    "query.versions.metadata.v1",
    "query.file.category.v1",
    "query.request.files.v1",
    "query.request.description.v1",
    "query.read.from.version.v1",
}

# Prefixo AMQP para extrair comando
ROUTING_KEY_PREFIX = "usiminas.req."

# ------------------------------------------------------------------ #
#  Diretórios e Log
# ------------------------------------------------------------------ #
TEMP_DIR = os.getenv("TEMP", os.getcwd())
LOG_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Log")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"consumer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

ALREADY_EXISTS_MARKERS = [
    "already exists", "ja existe", "já existe",
    "existente", "name already", "already active",
]
