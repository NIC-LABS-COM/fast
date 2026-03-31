"""
Configuracoes e constantes da aplicacao.
"""
import os

# OpenAI
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_MODEL = "gpt-4o"

# Diretorios
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_DIR = os.path.join(BASE_DIR, "Log")

# RabbitMQ
RABBITMQ_HOST = os.environ.get("RABBITMQ_HOST", "jackal.rmq.cloudamqp.com")
RABBITMQ_USER = os.environ.get("RABBITMQ_USER", "rhrstugr")
RABBITMQ_PASS = os.environ.get("RABBITMQ_PASS", "")
RABBITMQ_VHOST = os.environ.get("RABBITMQ_VHOST", "rhrstugr")
QUEUE_COMMANDS = "queue_vpn_usiminas"
QUEUE_RESPONSES = "queue_vpn_respostas"

# URLs dos scripts VBS no GitHub
VBS_URLS: dict[str, str] = {
    "criar_dominio":  "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/ScriptCriarDominio.vbs",
    "criar_elemento": "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/ScriptCriarElemento3.vbs",
    "criar_tabela":   "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/criarTabela.vbs",
    "criar_programa":  "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/ScriptCriarSE38.vbs",
    "editar_programa": "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/ScriptEditarSE38.vbs",
    "buscar_programa": "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/baixarArquivoSAP.vbs",
    "buscar_arquivo":  "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/buscaArquivo.vbs",
}

# SE38 - Tipos de programa
PROG_TYPE_OPTIONS: dict[str, str] = {
    "1": "Report",
    "M": "Module Pool",
    "I": "Include Program",
    "S": "Subroutine Pool",
    "F": "Function Group",
    "J": "Interface Pool",
    "K": "Class Pool",
    "T": "Type Pool",
}

# SE38 - Status do programa
PROG_STATUS_OPTIONS: dict[str, str] = {
    "T": "T - Programa de Teste",
    "P": "P - Programa Produtivo",
    "S": "S - Programa de Sistema",
}

# SE38 - Fluxo de execucao
PROG_FLOW_OPTIONS: dict[str, str] = {
    "Criar e Ativar": "Criar + Inserir Codigo + Ativar",
    "FIX":            "Editar Codigo Existente",
}

# Defaults de formulario
DEFAULT_REQUEST = ""
DEFAULT_PACKAGE = ""
DEFAULT_DOMAIN_DATATYPE = "CHAR"
DEFAULT_DOMAIN_LENGTH = "15"

# Opcoes de combobox
TABART_OPTIONS: dict[str, str] = {
    "APPL0": "Dados mestre, tabelas transparentes",
    "APPL1": "Dados de movimento, tabelas transparentes",
    "APPL2": "Organizacao e customizing",
}

TABKAT_OPTIONS: dict[str, str] = {
    "0": "0 ate 17.000",
    "1": "17.000 ate 68.000",
    "2": "68.000 ate 270.000",
    "3": "270.000 ate 1.000.000",
    "4": "1.000.000 ate 4.300.000",
    "5": "4.300.000 ate 8.700.000",
    "6": "8.700.000 ate 17.000.000",
    "7": "17.000.000 ate 34.000.000",
}

# Icones usados no log de respostas
STATUS_ICONS: dict[str, str] = {
    "sucesso": "[OK]",
    "ja_existe": "[!!]",
    "erro": "[ERRO]",
}
