"""Constantes e configuracao do Consumer RabbitMQ."""

import os
import sys
from datetime import datetime
from pathlib import Path

# Carrega .env da raiz do projeto
# Use ENV_FILE=.env.dev para apontar para o ambiente de dev
_project_root = Path(__file__).resolve().parent.parent
_env_file = os.environ.get("ENV_FILE", ".env")
_env_path = _project_root / _env_file
if _env_path.exists():
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _key, _val = _line.split("=", 1)
                _val = _val.strip().strip('"').strip("'")
                os.environ[_key.strip()] = _val  # sobrescreve para respeitar o arquivo escolhido

# ------------------------------------------------------------------ #
#  Configuracao RabbitMQ
# ------------------------------------------------------------------ #
BROKER_HOST     = os.environ.get("BROKER_HOST",     "rabbitmq.nic-labs.com")
BROKER_PORT     = int(os.environ.get("BROKER_PORT", "5671"))
BROKER_USER     = os.environ.get("BROKER_USER",     "nicaiuser")
BROKER_PASSWORD = os.environ.get("BROKER_PASSWORD", "USOYl8SM*L1P3iOXs@f3$oYLwdOA6xJv")
BROKER_VHOST    = os.environ.get("BROKER_VHOST",    "/")
BROKER_TIMEOUT  = int(os.environ.get("BROKER_TIMEOUT", "150000"))

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

# Diretório local com os scripts VBS empacotados junto ao executável
def _get_vbs_dir() -> str:
    if getattr(sys, "frozen", False):
        # Executável PyInstaller — arquivos extraídos em sys._MEIPASS
        return os.path.join(sys._MEIPASS, "vbs")
    # Desenvolvimento — pasta vbs/ na raiz do projeto
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "vbs")

VBS_DIR = _get_vbs_dir()

# Mapeamento: routing_key -> nome do arquivo VBS local
VBS_BY_ROUTING_KEY: dict[str, str] = {
    "command.fix.v1":             "ScriptEditarSE38.vbs",
    "command.revert.v1":          "ScriptEditarSE38.vbs",
    "command.create.v1":          "ScriptCriarSE38.vbs",
    "command.buscar.v1":          "buscarProgSap.vbs",
    "query.read.file.v1":         "baixarArquivoSAP.vbs",
    "query.requests.v1":          "buscarRequests.vbs",
    "query.all.files.v1":         "buscaReports.vbs",
    "query.all.packages.v1":      "buscaPacotes.vbs",
    "query.versions.metadata.v1": "buscaVersionsMetadata.vbs",
    "query.file.category.v1":     "buscaCategoryByFileName.vbs",
    "query.request.files.v1":     "buscaAbapFilesByRequest.vbs",
    "query.request.description.v1": "buscaRequestDescription.vbs",
    "query.read.from.version.v1": "buscaConteudoPorVersao.vbs",
    "query.search.v1":            "buscarObjetos.vbs",
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
    "query.search.v1",
}

# Prefixo AMQP para extrair comando
ROUTING_KEY_PREFIX = "usiminas.req."

# ------------------------------------------------------------------ #
#  Diretórios e Log
# ------------------------------------------------------------------ #
LOG_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Log")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"consumer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

ALREADY_EXISTS_MARKERS = [
    "already exists", "ja existe", "já existe",
    "existente", "name already", "already active",
]
