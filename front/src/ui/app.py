"""
Janela principal da aplicacao.
Orquestra servicos, abas e o painel de log.
"""
import logging
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Dict, List

from core.config import STATUS_ICONS
from services.ai_service import AIService
from services.publisher_service import PublisherService
from services.rabbitmq_service import RabbitMQService
from ui.tabs.dominio_tab import DominioTab
from ui.tabs.elemento_tab import ElementoTab
from ui.tabs.tabela_tab import TabelaTab
from ui.tabs.se38_tab import SE38Tab
from ui.tabs.queries_tab import QueriesTab
from ui.widgets.log_panel import LogPanel

logger = logging.getLogger("sap_publisher")


class SapPublisherApp:
    """Aplicacao principal SAP ABAP Dictionary Publisher."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("SAP ABAP Dictionary - Publisher")
        self.root.geometry("1450x920")

        self.status_var = tk.StringVar(value="Pronto")
        # chave = correlationId, valor = dict original da mensagem
        self._pending: Dict[str, Dict] = {}
        self._responses: List[Dict] = []

        self._rabbitmq  = RabbitMQService()
        self._publisher = PublisherService()
        self._ai        = AIService()

        self._build_ui()
        self._rabbitmq.start_listener(
            lambda resp: self.root.after(0, self._handle_response, resp)
        )
        logger.info("SapPublisherApp iniciado")

    # ------------------------------------------------------------------ #
    #  Status
    # ------------------------------------------------------------------ #
    def _set_status(self, text: str) -> None:
        self.status_var.set(text)
        self.root.update_idletasks()

    # ------------------------------------------------------------------ #
    #  Resposta do consumer
    # ------------------------------------------------------------------ #
    def _handle_response(self, response) -> None:
        # Formato Query: lista pura de objetos
        if isinstance(response, list):
            self._queries_tab.display_response(response)
            return

        # Formato Query: string simples (ex: query.file.category.v1)
        if isinstance(response, str):
            self._queries_tab.display_string_response(response)
            return

        correlation_id = response.get("correlationId", "")

        # Formato V1 Fix/Create: {hasError, errorMessage, isCodeError, correlationId}
        if "hasError" in response and "isCodeError" in response:
            correlation_id = response.get("correlationId", "")
            has_error = response.get("hasError", False)
            error_msg = response.get("errorMessage", "") or ""
            self._pending.pop(correlation_id, None)
            if has_error:
                self._log_panel.append(f"[ERRO] {error_msg}", "erro")
                self._set_status("Erro ao executar comando SAP.")
            else:
                self._log_panel.append("[OK] Comando executado com sucesso no SAP.", "sucesso")
                self._set_status("Concluido com sucesso.")
            return

        # Formato V1: {"content": "..."} ou {"error": "..."}
        if "content" in response:
            code = response["content"].replace("\\n", "\n")
            self._se38_tab.load_source_code(code)
            self._set_status("Codigo carregado do SAP.")
            self._pending.pop(correlation_id, None)
            return
        
        if "error" in response:
            error_msg = response["error"]
            self._log_panel.append(f"[ERRO] Falha ao buscar programa: {error_msg}", "erro")
            self._set_status("Erro ao buscar programa.")
            self._pending.pop(correlation_id, None)
            return
        
        # Resposta sem formato conhecido — ignora silenciosamente
        if "action" not in response:
            logger.warning("Resposta ignorada (formato desconhecido): %s", response)
            self._pending.pop(correlation_id, None)
            return

        # Formato legado
        action = response["action"]
        obj = response.get("object_name", "?")
        status = response.get("status", "?")
        message = response.get("message", "")
        icon = STATUS_ICONS.get(status, "[?]")

        self._pending.pop(correlation_id, None)

        # Resposta de buscar_programa — popula editor, nao exibe resumo
        if action == "buscar_programa":
            if status == "sucesso":
                code = message.replace("\\n", "\n")
                self._se38_tab.load_source_code(code)
                self._set_status("Codigo carregado do SAP.")
            else:
                self._log_panel.append(f"[ERRO] Falha ao buscar programa: {message}", "erro")
                self._set_status("Erro ao buscar programa.")
            return

        # Resposta de buscar_arquivo — popula editor com conteudo do TXT
        if action == "buscar_arquivo":
            if status == "sucesso":
                content = message.replace("\\n", "\n")
                self._se38_tab.load_file_content(content)
                self._set_status("Arquivo carregado.")
            else:
                self._log_panel.append(f"[ERRO] Falha ao buscar arquivo: {message}", "erro")
                self._set_status("Erro ao buscar arquivo.")
            return

        self._log_panel.append(f"{icon} [{correlation_id[:8]}] {action} | {obj} | {message}", status)
        self._responses.append(response)

        restantes = len(self._pending)
        if restantes > 0:
            self._set_status(f"Processando... ({restantes} restante(s))")
        else:
            self._show_summary()

    def _show_summary(self) -> None:
        if not self._responses:
            return
        total = len(self._responses)
        ok = sum(1 for r in self._responses if r.get("status") == "sucesso")
        exists = sum(1 for r in self._responses if r.get("status") == "ja_existe")
        errs = sum(1 for r in self._responses if r.get("status") == "erro")

        summary = (
            f"Resumo do processamento ({total} operacao(es)):\n\n"
            f"  Sucesso: {ok}\n  Ja existia: {exists}\n  Erro: {errs}\n\n"
        )
        if errs > 0:
            summary += "Detalhes dos erros:\n"
            for r in self._responses:
                if r.get("status") == "erro":
                    summary += f"  - {r.get('action')} | {r.get('object_name')}: {r.get('message', '')}\n"

        self._set_status(f"Concluido: {ok} ok, {exists} ja existiam, {errs} erro(s)")
        self._responses.clear()

        if errs > 0:
            messagebox.showwarning("Concluido com erros", summary)
        else:
            messagebox.showinfo("Concluido", summary)

    # ------------------------------------------------------------------ #
    #  Callback de publicacao (chamado pelas abas)
    # ------------------------------------------------------------------ #
    def _on_publish(self, messages: List[Dict]) -> None:
        self._pending.clear()
        self._responses.clear()

        total = len(messages)

        # Feedback imediato antes da conexao bloqueante
        self._set_status(f"Conectando e enviando {total} mensagem(ns) sequencialmente...")
        self._log_panel.append(f"Conectando ao RabbitMQ... ({total} etapas)")
        self.root.update_idletasks()

        if total == 1:
            ok = self._rabbitmq.publish(messages[0])
            if ok:
                self._pending[messages[0]["correlationId"]] = messages[0]
                self._set_status("Enviado. Aguardando resposta...")
                self._log_panel.append(
                    f"Enviado [{messages[0]['correlationId'][:8]}] — aguardando consumer..."
                )
            else:
                self._set_status("Falha ao enviar")
                messagebox.showerror("Erro", "Falha ao enviar para fila.")
            return

        # Multiplas mensagens: executa cadeia sequencial em thread separada
        def _chain_worker():
            def on_step(idx, total, response):
                action = response.get("action", "?")
                obj = response.get("object_name", "?")
                status = response.get("status", "?")
                icon = STATUS_ICONS.get(status, "[?]")
                self.root.after(0, self._log_panel.append,
                    f"{icon} [{idx+1}/{total}] {action} | {obj} | {response.get('message', '')}", status)
                self.root.after(0, self._set_status,
                    f"Processando etapa {idx+1}/{total}: {action} | {obj} → {status}")

            responses = self._rabbitmq.publish_chain(
                messages, timeout=120.0, on_step=on_step,
            )
            self.root.after(0, self._handle_chain_complete, messages, responses)

        thread = threading.Thread(target=_chain_worker, daemon=True)
        thread.start()

    def _handle_chain_complete(self, messages: List[Dict], responses: List[Dict]) -> None:
        """Exibe resumo final apos cadeia sequencial terminar."""
        total = len(messages)
        processed = len(responses)
        ok = sum(1 for r in responses if r.get("status") == "sucesso")
        exists = sum(1 for r in responses if r.get("status") == "ja_existe")
        errs = sum(1 for r in responses if r.get("status") == "erro")
        skipped = total - processed

        summary = (
            f"Resumo do processamento ({total} etapa(s)):\n\n"
            f"  Sucesso: {ok}\n  Ja existia: {exists}\n  Erro: {errs}\n"
        )
        if skipped > 0:
            summary += f"  Nao executadas (abortadas): {skipped}\n"
        summary += "\n"

        if errs > 0:
            summary += "Detalhes dos erros:\n"
            for r in responses:
                if r.get("status") == "erro":
                    summary += f"  - {r.get('action')} | {r.get('object_name')}: {r.get('message', '')}\n"
            if skipped > 0:
                summary += f"\n{skipped} etapa(s) subsequente(s) foram canceladas.\n"

        self._set_status(f"Concluido: {ok} ok, {exists} ja existiam, {errs} erro(s), {skipped} canceladas")

        if errs > 0:
            messagebox.showwarning("Concluido com erros", summary)
        else:
            messagebox.showinfo("Concluido", summary)

    def _on_publish_v1(self, v1_data: Dict) -> None:
        """Publica mensagem V1 usando exchange topic."""
        self._pending.clear()
        self._responses.clear()

        # Feedback imediato
        self._set_status("Conectando ao RabbitMQ (V1)...")
        self._log_panel.append(f"Conectando ao RabbitMQ...")
        self.root.update_idletasks()

        ok = self._rabbitmq.publish_v1(
            routing_key=v1_data["routing_key"],
            payload=v1_data["payload"],
        )
        if ok:
            self._pending[v1_data["correlation_id"]] = v1_data["payload"]
            self._set_status("Enviado V1. Aguardando resposta...")
            self._log_panel.append(
                f"Enviado [{v1_data['correlation_id'][:8]}] — aguardando consumer..."
            )
        else:
            self._set_status("Falha ao enviar V1")
            messagebox.showerror("Erro", "Falha ao enviar para exchange.")

    # ------------------------------------------------------------------ #
    #  UI
    # ------------------------------------------------------------------ #
    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(main)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        # log_fn e set_status_fn sao lambdas que fecham sobre self.
        # _log_panel ainda nao existe aqui, mas so sera chamada apos o build.
        def log_fn(text: str, status: str = "") -> None:
            self._log_panel.append(text, status)

        dom_tab = DominioTab(notebook, self._publisher, self._on_publish, log_fn)
        notebook.add(dom_tab, text="  Criar Dominio  ")

        elem_tab = ElementoTab(notebook, self._publisher, self._on_publish, log_fn)
        notebook.add(elem_tab, text="  Criar Elemento  ")

        tab_tab = TabelaTab(
            notebook, self._publisher, self._ai,
            self._on_publish, log_fn, self._set_status,
        )
        notebook.add(tab_tab, text="  Criar Tabela  ")

        self._se38_tab = SE38Tab(
            notebook, self._publisher, self._ai,
            self._on_publish, self._on_publish_v1, log_fn, self._set_status,
        )
        notebook.add(self._se38_tab, text="  SE38 - Programa ABAP  ")

        self._queries_tab = QueriesTab(
            notebook, self._on_publish_v1, log_fn, self._set_status,
        )
        notebook.add(self._queries_tab, text="  Queries (Requests / Reports)  ")

        self._log_panel = LogPanel(main)
        self._log_panel.pack(fill=tk.X, pady=(4, 4))

        bottom = ttk.Frame(main)
        bottom.pack(fill=tk.X)
        ttk.Button(bottom, text="Limpar Log", command=self._log_panel.clear).pack(side=tk.RIGHT, padx=4)
        ttk.Label(bottom, textvariable=self.status_var, anchor="w").pack(fill=tk.X, side=tk.LEFT)
