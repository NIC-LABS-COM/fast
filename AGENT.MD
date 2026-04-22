# AGENT.md — Consumer RabbitMQ (Usiminas / SAP GUI)

> Guia obrigatório para qualquer agente (Claude / Copilot / humano) que for editar este repositório.
> O objetivo é garantir **zero perda de mensagem**, **alta resiliência** e **código previsível** em um ambiente que integra RabbitMQ + SAP GUI via VBScript.

---

## 1. Visão Geral do Projeto

- **Tipo:** Python Queue Consumer (pacote `consumer/`)
- **Broker:** RabbitMQ via `pika` (BlockingConnection, TLS/SSL porta 5671)
- **Destino das mensagens:** Scripts `.vbs` executados via `cscript.exe` que automatizam o **SAP GUI** (transações SE38, SE80, SE09, SE11 etc.)
- **Empacotamento:** PyInstaller → `ConsumerDev.exe` (ambiente dev) e `ConsumerTT.exe` (produção)
- **Persistência de dados:** **NÃO há DB aplicacional**. O "estado" vive no SAP e nos arquivos `.vbs`. Logs em `consumer/Log/`.
- **Concurrency model:** `pika` single-thread loop + `basic_qos(prefetch_count=1)`. NÃO introduzir `threading` sem antes revisar a [secção 6 — Concurrency].

### Duas arquiteturas coexistindo
1. **Legado** → fila `queue_vpn_usiminas`, payload `{action, vbs_url, args, correlationId}`, resposta em `queue_vpn_respostas`. Handler em [consumer/legacy.py](consumer/legacy.py).
2. **V1 (event-driven)** → exchange topic `x.to-client.topic`, binding `usiminas.req.#.v1`, fila `q.usiminas.v1`. Routing key dita qual VBS rodar. Handler em [consumer/v1.py](consumer/v1.py).

Toda feature nova **deve** ir para a arquitetura **V1**. Legado está em manutenção mínima.

---

## 2. Layout do Repositório

```
fast/
├── ConsumerDev.py / ConsumerTT.py     # Entrypoints PyInstaller (setam ENV_FILE)
├── consumer/                           # Pacote principal
│   ├── main.py                         # Bootstrap: declara exchange/filas/bindings
│   ├── config.py                       # Carrega .env, constantes, VBS_BY_ROUTING_KEY
│   ├── connection.py                   # pika + SSL + retry exponencial de conexão
│   ├── legacy.py                       # Handler da fila legado
│   ├── v1.py                           # Handler V1 (~700 linhas, várias publish_*)
│   ├── parsers.py                      # Parsers de TXT/JSON vindos dos VBS
│   ├── vbs.py                          # get_vbs_path + execute_vbs (subprocess cscript)
│   ├── retry.py                        # Decorator with_retry_and_ack (ACK/NACK/DLQ)
│   ├── errors.py                       # TransientError / PermanentError
│   ├── logger.py                       # log() simples com arquivo datado
│   └── Log/                            # consumer_YYYYMMDD_HHMMSS.txt (gitignored)
├── vbs/                                # Scripts SAP GUI (empacotados no .exe via PyInstaller)
├── reports/                            # Reports ABAP de referência (fonte do que os VBS buscam)
├── front/                              # Frontend auxiliar (fora do escopo do consumer)
└── dist/ConsumerTT.exe                 # Binário versionado (única exceção no .gitignore)
```

---

## 3. Setup & Comandos

```bash
# Instalar dependências (criar requirements.txt caso não exista — hoje o projeto só depende de 'pika')
pip install pika

# Rodar local apontando para .env.dev
python ConsumerDev.py

# Rodar produção (lê .env)
python ConsumerTT.py

# Empacotar
pyinstaller ConsumerDev.spec       # gera dist/ConsumerDev.exe
pyinstaller ConsumerTT.spec        # gera dist/ConsumerTT.exe

# Testes (quando adicionados)
pytest -v
```

### Variáveis de ambiente (ver [consumer/config.py](consumer/config.py))
- `ENV_FILE` (default `.env`) — permite `ENV_FILE=.env.dev` para trocar ambiente
- `BROKER_HOST`, `BROKER_PORT`, `BROKER_USER`, `BROKER_PASSWORD`, `BROKER_VHOST`
- `BROKER_TIMEOUT` (ms, default 150000)
- `MAX_RETRY_ATTEMPTS` (default 3), `RETRY_BACKOFF_SECONDS` (default 2)
- `MAX_RECONNECT_ATTEMPTS` (default 10), `RECONNECT_BACKOFF_SECONDS` (default 5)

> **Nunca** commitar `.env*`. O `.gitignore` já cobre, mas confirme antes de `git add`.
>
> **Credenciais default em [consumer/config.py](consumer/config.py) são intencionais:** o binário PyInstaller roda em máquinas que não têm `.env`, então os valores precisam estar embutidos. O `.env` só entra em cena para sobrescrever em ambiente dev.

---

## 4. Contrato de Mensagens (Arquitetura V1)

### Entrada — `x.to-client.topic` com routing key `usiminas.req.<comando>.v1`

Routing keys conhecidas estão em `VBS_BY_ROUTING_KEY` em [consumer/config.py:73](consumer/config.py#L73). Resumo:

| Routing key                         | Tipo    | VBS mapeado                     |
| ----------------------------------- | ------- | ------------------------------- |
| `command.fix.v1` / `command.revert.v1` | command | `ScriptEditarSE38.vbs`          |
| `command.create.v1`                 | command | `ScriptCriarSE38.vbs`           |
| `command.buscar.v1`                 | command | `buscarProgSap.vbs`             |
| `query.read.file.v1`                | query   | `baixarArquivoSAP.vbs`          |
| `query.requests.v1`                 | query   | `buscarRequests.vbs`            |
| `query.all.files.v1`                | query   | `buscaReports.vbs`              |
| `query.all.packages.v1`             | query   | `buscaPacotes.vbs`              |
| `query.versions.metadata.v1`        | query   | `buscaVersionsMetadata.vbs`     |
| `query.file.category.v1`            | query   | `buscaCategoryByFileName.vbs`   |
| `query.request.files.v1`            | query   | `buscaAbapFilesByRequest.vbs`   |
| `query.request.description.v1`      | query   | `buscaRequestDescription.vbs`   |
| `query.read.from.version.v1`        | query   | `buscaConteudoPorVersao.vbs`    |
| `query.search.v1`                   | query   | `buscarObjetos.vbs`             |

**Ao adicionar nova routing key:** atualizar `VBS_BY_ROUTING_KEY` **e** (se for query) `QUERY_ROUTING_KEYS`.

### Saída
- **Commands** → `publish_fix_response` → `{hasError, errorMessage, isCodeError, correlationId}`
- **Queries** (lista)  → `publish_query_response` → JSON array puro
- **Queries** (string) → `publish_string_response` → string JSON
- **read.file**        → `publish_read_file_response` → `{fileName, content}` ou `{error}`
- Respostas vão para `properties.reply_to` (RPC pattern) com `correlation_id` espelhado.

---

## 5. Mensageria — Regras Inegociáveis

### 5.1 Zero perda silenciosa
- **Nunca** use `except Exception: pass`.
- Todo erro é logado **e** classificado como transitório ou permanente.

### 5.2 Classificação de falhas
| Situação                                          | Classe              | Ação         |
| ------------------------------------------------- | ------------------- | ------------ |
| Payload malformado, routing key desconhecida, VBS ausente | `PermanentError`    | NACK→DLQ direto |
| Timeout SAP GUI, sessão ocupada, falha de rede    | `TransientError` (ou exceção genérica) | Retry com backoff → se esgotar, DLQ |
| Negócio irrecuperável (ex.: objeto já existe e regra proíbe) | `PermanentError` | NACK→DLQ |
| Sucesso                                           | —                   | ACK manual |

Definições em [consumer/errors.py](consumer/errors.py).

### 5.3 ACK manual obrigatório
- **Todos** os consumers (V1 e legado) já operam com `auto_ack=False` + decorator `@with_retry_and_ack()` de [consumer/retry.py](consumer/retry.py).
- Handlers novos **devem** seguir o mesmo padrão — deixar exceções propagarem para o decorator, que cuida do ACK/NACK/retry.
- **Nunca** registrar `basic_consume` com `auto_ack=True`.

### 5.4 DLQ
- Exchange DLX: `x.to-client.dlx`
- Fila DLQ: `q.usiminas.v1.dlq`
- Após esgotar `MAX_RETRY_ATTEMPTS` → `basic_nack(requeue=False)` → mensagem vai para DLQ.
- **Nunca** `requeue=True` em NACK — isso cria loops infinitos.

### 5.5 Reconexão
- Use sempre `get_rabbitmq_connection()` de [consumer/connection.py](consumer/connection.py) — ele já faz backoff exponencial capped em 60s e até `MAX_RECONNECT_ATTEMPTS` tentativas.

---

## 6. Concurrency

O consumer roda **single-threaded** por desenho (`pika.BlockingConnection` não é thread-safe para publish/ack).

- **NÃO** chame `basic_publish` / `basic_ack` / `basic_nack` de uma thread diferente da que está rodando `start_consuming()`.
- Se precisar paralelizar trabalho pesado (ex.: múltiplas sessões SAP), o padrão correto é:
  1. Submeter o trabalho a uma thread worker (ou `concurrent.futures`).
  2. Quando o worker terminar, agendar o ACK/publish via `connection.add_callback_threadsafe(callable)`.
- Nunca compartilhar o mesmo `channel` entre threads.
- Execução do VBS via `subprocess.run` é síncrona e bloqueante — isso é intencional, pois o SAP GUI é um recurso single-session por usuário Windows.

---

## 7. Execução de VBS (SAP GUI)

- Toda chamada ao SAP GUI passa por [consumer/vbs.py](consumer/vbs.py) → `execute_vbs(path, args)`.
- Usa `cscript.exe //nologo` (path montado via `%SYSTEMROOT%`).
- stdout decodificado como `oem`, stderr como `cp1252` — não alterar sem motivo forte (vieram de debugging real de encoding SAP/Windows).
- Prefixo `SAP_ERROR:` em stderr é removido antes de retornar para o chamador.
- Em desenvolvimento, `VBS_DIR` aponta para `vbs/` na raiz; no `.exe`, aponta para `sys._MEIPASS/vbs` (PyInstaller). Isso é tratado em [consumer/config.py:63](consumer/config.py#L63).
- Para adicionar um VBS novo:
  1. Colocar o `.vbs` em `vbs/`.
  2. Registrar em `VBS_BY_ROUTING_KEY`.
  3. Rebuildar o `.spec` — o PyInstaller já inclui `datas=[('vbs', 'vbs')]`.

---

## 8. Estilo de Código

- **PEP 8** estrito.
- **Type hints obrigatórios** em toda função pública: `def process(ch, body: bytes) -> None:`.
- **Docstrings Google Style** apenas em funções com lógica não-trivial (evitar docstrings que só repetem o nome da função).
- Português é aceito em mensagens de log e comentários de domínio (alinhado ao padrão existente); nomes de símbolos sempre em inglês.
- **Não adicionar comentários óbvios.** Só comente o *porquê*, nunca o *o quê*.
- Preferir **pure functions** em `parsers.py`; efeitos colaterais (publish, log, subprocess) ficam nos handlers.
- Imports relativos dentro do pacote: `from .config import ...`.

### O que evitar
- ❌ `except Exception: pass`
- ❌ `auto_ack=True` em handlers novos
- ❌ `print(...)` fora do `logger.log()` (exceto banner de boot do `main.py`)
- ❌ Introduzir dependências pesadas (SQLAlchemy, FastAPI, Celery) sem discutir — o projeto é deliberadamente minimalista
- ❌ Remover os defaults de credenciais em [consumer/config.py:25-30](consumer/config.py#L25-L30) — **isso é intencional**: o consumer é empacotado como `.exe` standalone e precisa conectar sem depender de `.env` externo. O `.env` apenas sobrescreve quando presente (ambiente dev).

---

## 9. Logging

- Sempre via `log()` de [consumer/logger.py](consumer/logger.py).
- Prefixos convencionais: `[CONN]`, `[RETRY]`, `[V1]`, `[LEGACY]` — manter o padrão ao adicionar módulos novos.
- Nunca logar payload completo com credenciais / tokens.
- Arquivo novo por execução: `consumer/Log/consumer_YYYYMMDD_HHMMSS.txt`. Há limpeza manual; ao escrever ferramentas de retenção, preservar os 30 últimos dias.

---

## 10. Mudanças Seguras (checklist antes de abrir PR)

- [ ] Rodei `python ConsumerDev.py` localmente e o boot imprimiu `[CONN] Conectado com sucesso`.
- [ ] Se alterei routing keys: atualizei `VBS_BY_ROUTING_KEY` **e** `QUERY_ROUTING_KEYS` quando aplicável.
- [ ] Se alterei handler: verifiquei que falhas resultam em `PermanentError` ou `TransientError` — nunca exceção silenciosa.
- [ ] Se alterei `main.py` / `connection.py`: testei a reconexão derrubando o broker.
- [ ] Não quebrei o build: `pyinstaller ConsumerDev.spec` roda sem warnings.
- [ ] Nenhum `.env` ou credencial nova foi commitada.

---

## 11. Objetivo do Agente (resumo operacional)

Atue como **especialista em sistemas distribuídos com Python + integração SAP**, priorizando:

1. **Durabilidade** — toda mensagem tem destino determinístico (ACK, retry ou DLQ).
2. **Observabilidade** — cada decisão aparece no log com contexto suficiente para debug remoto.
3. **Simplicidade** — preferir 20 linhas legíveis a uma abstração elegante. Este projeto roda em um Windows ligado 24/7; a próxima pessoa a debugar é alguém sob pressão.
4. **Fidelidade à arquitetura existente** — novas features vão para V1, não para legado.
5. **Respeito ao SAP** — sessão GUI é frágil; nunca paralelizar chamadas que compartilham o mesmo usuário SAP.

Quando em dúvida entre "fazer genérico" e "fazer direto", **faça direto** e documente o porquê.
