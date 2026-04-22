# Prompt padrão para gerar AGENT.md em outros projetos

Este arquivo contém o prompt reutilizável que você cola em qualquer outro projeto (front, back, scripts) para que o agente (Claude / Copilot / etc.) construa um `AGENT.md` específico daquele repo, lendo o código real antes de escrever.

Uso:
1. Abra o projeto alvo no Claude Code / Copilot.
2. Copie tudo dentro do bloco abaixo e envie como mensagem.
3. Revise o AGENT.md gerado antes de commitar.

---

## Prompt (copiar daqui até o fim do bloco)

````
Preciso que você crie um AGENT.md completo para este projeto.
Não copie template genérico — leia o código de verdade antes de escrever.

## Passo 1 — Explorar antes de escrever

Antes de gerar qualquer linha do AGENT.md, você DEVE:
1. Listar a estrutura de pastas e identificar o stack real (framework, linguagem, build tool, gerenciador de pacotes).
2. Abrir os arquivos de entrada/bootstrap (ex.: main.py, index.ts, App.tsx, server.js).
3. Abrir os arquivos de config (package.json, requirements.txt, pyproject.toml, tsconfig, .env.example, docker-compose).
4. Identificar padrões recorrentes: como erros são tratados, como rotas/handlers/componentes são organizados, onde mora a lógica de negócio.
5. Identificar contratos: rotas HTTP, schemas, eventos, payloads, nomes de filas.
6. Me listar em 1 parágrafo o que você descobriu antes de escrever o arquivo.

Se o projeto tiver mais de uma arquitetura (legado + nova, v1 + v2), documente AS DUAS e diga explicitamente qual é o caminho recomendado para novas features.

## Passo 2 — Seções obrigatórias do AGENT.md

1. **Visão Geral do Projeto** — tipo, stack, propósito em 1 parágrafo, modelo de deploy/distribuição.
2. **Layout do Repositório** — árvore comentada explicando o que cada pasta/arquivo principal faz. Use links markdown relativos para os caminhos.
3. **Setup & Comandos** — install, run dev, run prod, test, build. Só comandos que funcionam de verdade neste projeto (verifique nos scripts do package.json / Makefile / etc.).
4. **Variáveis de Ambiente** — liste as que o código realmente lê, com defaults quando houver. Reforce o que NÃO pode ser commitado.
5. **Contratos** — APIs, rotas, eventos, schemas, props de componentes reutilizáveis. Tabelas sempre que fizer sentido. Inclua caminhos com linha (arquivo.ext:linha) apontando para onde estão declarados.
6. **Regras Inegociáveis** — o que nunca pode ser feito (ex.: `except: pass`, rotas sem auth, mutação direta de estado, queries sem índice). Cada regra com um "porquê".
7. **Classificação de erros** (se o projeto lida com integrações/filas/API externa) — o que é transitório, o que é definitivo, como cada um é tratado.
8. **Estilo de Código** — linter, formatador, convenções de nomenclatura, idioma dos comentários. Cite as configs existentes (eslint, ruff, black, prettier).
9. **O que evitar** — lista de antipadrões específicos deste projeto, não genéricos.
10. **Logging / Observabilidade** — como logar, prefixos convencionais, o que NUNCA pode aparecer em log.
11. **Checklist antes de abrir PR** — itens acionáveis, não platitudes.
12. **Objetivo do Agente** — 3 a 5 princípios que guiam qualquer decisão (ex.: "simplicidade > genericidade", "fidelidade à arquitetura existente").

## Passo 3 — Princípios de qualidade do AGENT.md

- **Específico, não genérico.** "Use type hints" é fraco. "Type hints obrigatórios em funções públicas; `dict[str, Any]` proibido — usar TypedDict" é forte.
- **Com endereço.** Toda regra que referencia código deve vir com [arquivo.ext](caminho/arquivo.ext) ou [arquivo.ext:linha](caminho/arquivo.ext#Llinha).
- **Assumir leitor sob pressão.** A próxima pessoa a ler esse doc está debugando às 23h. Escreva pra ela.
- **Explicar o porquê** de regras não-óbvias. Regra sem motivo é ignorada.
- **Marcar débito técnico** quando encontrar — não esconda. Ex.: "atenção: arquivo X ainda usa padrão antigo, migrar antes de tocar nele".
- **Sem encher linguiça.** Se uma seção não se aplica ao projeto, omita em vez de escrever "N/A".
- **Idioma:** escreva no idioma que os comentários e commits do projeto usam hoje. Não force inglês se o time trabalha em português.

## Passo 4 — Não invente

- Se o projeto NÃO tem testes, diga "testes ausentes — adicionar antes de crescer" em vez de inventar um comando pytest.
- Se o projeto NÃO tem CI, não finja que tem.
- Se uma dependência "comum" (typescript, docker, redis) não existe aqui, não liste.
- Dependências, scripts e configs só entram no AGENT.md se existirem de verdade no repo.

Agora prossiga: explore, me resuma o que encontrou em 1 parágrafo, e depois gere o AGENT.md no arquivo AGENT.md na raiz.
````

---

## Adendos opcionais por tipo de projeto

Cole o adendo relevante logo abaixo do prompt principal, dentro do mesmo bloco, antes da última linha ("Agora prossiga...").

### Frontend (React / Vue / Next / Svelte)
```
Observações extras para front:
- Documente o gerenciamento de estado (Redux, Zustand, Context, Pinia, stores).
- Documente a biblioteca de UI (Material, shadcn, Tailwind, CSS Modules) e as regras de uso.
- Documente a camada de chamada a API (axios wrapper, fetch helper, hooks de data fetching como React Query/SWR).
- Documente o sistema de rotas (react-router, file-based routing do Next).
- Inclua a convenção de nome de componentes, pasta de atoms/molecules se houver, e como criar um componente novo.
- Liste os endpoints/filas consumidos — se o backend está em outro repo, referencie o AGENT.md de lá.
```

### Backend (API REST / GraphQL)
```
Observações extras para back:
- Documente cada rota/resolver: método, path, body, response, auth requerida, status codes esperados.
- Documente o modelo de autenticação/autorização (JWT, sessão, API key, RBAC).
- Documente a camada de persistência: ORM, migrations, como criar uma migration nova, convenção de nomes de tabela.
- Documente middlewares obrigatórios (logging, rate limit, CORS, error handler).
- Documente o contrato com filas/eventos se houver (publisher/consumer patterns).
- Documente o pipeline de CI/CD e o ambiente de deploy (Docker, k8s, serverless).
```

### Consumer / Worker (filas, scheduled jobs)
```
Observações extras para consumer:
- Documente cada fila/tópico consumido, payload esperado e routing keys.
- Documente a estratégia de ACK, retry, DLQ.
- Documente como distinguir erro transitório vs definitivo.
- Documente o modelo de concorrência (single-thread, pool, async).
- Documente reconexão ao broker em caso de queda.
```

---

## Dicas de uso

- Se o projeto já tem um AGENT.md antigo, adicione no prompt: *"Já existe um AGENT.md na raiz — revise, mantenha o que ainda é verdade e atualize/remova o que estiver desatualizado."*
- Se o projeto tem múltiplos sub-pacotes (monorepo), peça um AGENT.md na raiz + um por pacote, com o da raiz apontando para os filhos.
- Depois de gerado, leia o AGENT.md inteiro. Se alguma regra não fizer sentido, remova — melhor ter um doc curto e fiel do que longo e mentiroso.
