"""
Aba Queries — Sub-abas para cada tipo de consulta SAP.
Cada sub-aba tem seus inputs, botao e painel de resultado proprio.
"""
import json
import uuid
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Dict


def _make_result_panel(parent) -> tk.Text:
    """Cria painel de resultado padrao (dark theme) com scrollbar."""
    frame = ttk.LabelFrame(parent, text="Retorno da Query")
    frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

    text = tk.Text(
        frame, wrap=tk.WORD, font=("Consolas", 10),
        bg="#1e1e1e", fg="#d4d4d4", insertbackground="#d4d4d4",
    )
    text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8, side=tk.LEFT)

    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text.yview)
    scrollbar.pack(fill=tk.Y, side=tk.RIGHT, padx=(0, 8), pady=8)
    text.configure(yscrollcommand=scrollbar.set)
    return text


class QueriesTab(ttk.Frame):
    """Container principal — Notebook interno com sub-abas de queries."""

    def __init__(
        self,
        parent,
        on_publish_v1: Callable[[Dict], None],
        log_fn: Callable[[str], None],
        set_status_fn: Callable[[str], None],
        **kwargs,
    ) -> None:
        super().__init__(parent, padding=4, **kwargs)
        self._on_publish_v1 = on_publish_v1
        self._log_fn = log_fn
        self._set_status = set_status_fn
        self._result_panels: list[tk.Text] = []
        self._build()

    # ------------------------------------------------------------------ #
    #  Build
    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill=tk.BOTH, expand=True)

        self._build_requests_tab()
        self._build_reports_tab()
        self._build_packages_tab()
        self._build_versions_tab()
        self._build_category_tab()
        self._build_request_files_tab()
        self._build_request_description_tab()
        self._build_read_from_version_tab()

    # ---- Requests ----
    def _build_requests_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  Requests  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        self._btn_requests = ttk.Button(
            top, text="Buscar Requests", command=self._query_requests, width=22,
        )
        self._btn_requests.pack(side=tk.LEFT)

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._requests_result)).pack(side=tk.RIGHT)

        self._requests_result = _make_result_panel(tab)
        self._result_panels.append(self._requests_result)

    # ---- Reports ----
    def _build_reports_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  Reports  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        self._btn_reports = ttk.Button(
            top, text="Buscar Reports", command=self._query_reports, width=22,
        )
        self._btn_reports.pack(side=tk.LEFT)

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._reports_result)).pack(side=tk.RIGHT)

        self._reports_result = _make_result_panel(tab)
        self._result_panels.append(self._reports_result)

    # ---- Pacotes ----
    def _build_packages_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  Pacotes  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        self._btn_packages = ttk.Button(
            top, text="Buscar Pacotes", command=self._query_packages, width=22,
        )
        self._btn_packages.pack(side=tk.LEFT)

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._packages_result)).pack(side=tk.RIGHT)

        self._packages_result = _make_result_panel(tab)
        self._result_panels.append(self._packages_result)

    # ---- Versions Metadata ----
    def _build_versions_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  Versions Metadata  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        ttk.Label(top, text="fileName:").pack(side=tk.LEFT)
        self._ver_filename_var = tk.StringVar()
        ttk.Entry(top, textvariable=self._ver_filename_var, width=30).pack(side=tk.LEFT, padx=(4, 12))

        ttk.Label(top, text="category:").pack(side=tk.LEFT)
        self._ver_category_var = tk.StringVar(value="PROGRAM")
        ttk.Combobox(
            top, textvariable=self._ver_category_var, width=20,
            values=["PROGRAM", "FUNCTION_MODULE", "CLASS"], state="readonly",
        ).pack(side=tk.LEFT, padx=(4, 12))

        self._btn_versions = ttk.Button(
            top, text="Buscar Versões", command=self._query_versions_metadata, width=18,
        )
        self._btn_versions.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._versions_result)).pack(side=tk.RIGHT)

        self._versions_result = _make_result_panel(tab)
        self._result_panels.append(self._versions_result)

    # ---- File Category ----
    def _build_category_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  File Category  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        ttk.Label(top, text="fileName:").pack(side=tk.LEFT)
        self._cat_filename_var = tk.StringVar()
        ttk.Entry(top, textvariable=self._cat_filename_var, width=30).pack(side=tk.LEFT, padx=(4, 12))

        self._btn_category = ttk.Button(
            top, text="Buscar Categoria", command=self._query_file_category, width=18,
        )
        self._btn_category.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._category_result)).pack(side=tk.RIGHT)

        self._category_result = _make_result_panel(tab)
        self._result_panels.append(self._category_result)

    # ---- Files by Request ----
    def _build_request_files_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  Files by Request  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Requests (separadas por vírgula):").pack(side=tk.LEFT)
        self._reqfiles_var = tk.StringVar()
        ttk.Entry(top, textvariable=self._reqfiles_var, width=50).pack(side=tk.LEFT, padx=(4, 12))

        self._btn_request_files = ttk.Button(
            top, text="Buscar Arquivos", command=self._query_request_files, width=18,
        )
        self._btn_request_files.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._reqfiles_result)).pack(side=tk.RIGHT)

        self._reqfiles_result = _make_result_panel(tab)
        self._result_panels.append(self._reqfiles_result)

    # ---- Request Description ----
    def _build_request_description_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  Request Description  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Request ID:").pack(side=tk.LEFT)
        self._reqdesc_var = tk.StringVar()
        ttk.Entry(top, textvariable=self._reqdesc_var, width=30).pack(side=tk.LEFT, padx=(4, 12))

        self._btn_request_desc = ttk.Button(
            top, text="Buscar Descrição", command=self._query_request_description, width=18,
        )
        self._btn_request_desc.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._reqdesc_result)).pack(side=tk.RIGHT)

        self._reqdesc_result = _make_result_panel(tab)
        self._result_panels.append(self._reqdesc_result)

    # ---- Read From Version ----
    def _build_read_from_version_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(tab, text="  Read From Version  ")

        top = ttk.Frame(tab)
        top.pack(fill=tk.X)

        ttk.Label(top, text="fileName:").pack(side=tk.LEFT)
        self._rfv_filename_var = tk.StringVar()
        ttk.Entry(top, textvariable=self._rfv_filename_var, width=30).pack(side=tk.LEFT, padx=(4, 12))

        ttk.Label(top, text="category:").pack(side=tk.LEFT)
        self._rfv_category_var = tk.StringVar(value="PROGRAM")
        ttk.Combobox(
            top, textvariable=self._rfv_category_var, width=20,
            values=["PROGRAM", "FUNCTION_MODULE", "CLASS"], state="readonly",
        ).pack(side=tk.LEFT, padx=(4, 12))

        ttk.Label(top, text="versionId:").pack(side=tk.LEFT)
        self._rfv_version_var = tk.StringVar()
        ttk.Entry(top, textvariable=self._rfv_version_var, width=20).pack(side=tk.LEFT, padx=(4, 12))

        self._btn_read_version = ttk.Button(
            top, text="Buscar Conteúdo", command=self._query_read_from_version, width=18,
        )
        self._btn_read_version.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(top, text="Limpar", width=10,
                   command=lambda: self._clear_panel(self._rfv_result)).pack(side=tk.RIGHT)

        self._rfv_result = _make_result_panel(tab)
        self._result_panels.append(self._rfv_result)

    # ------------------------------------------------------------------ #
    #  Acoes
    # ------------------------------------------------------------------ #
    def _query_requests(self) -> None:
        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "QueryFilterString": "",
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.requests.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn("Enviando query.requests.v1...")
        self._set_status("Buscando requests do SAP...")
        self._btn_requests.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_requests.configure(state=tk.NORMAL)

    def _query_reports(self) -> None:
        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.all.files.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn("Enviando query.all.files.v1...")
        self._set_status("Buscando reports do SAP...")
        self._btn_reports.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_reports.configure(state=tk.NORMAL)

    def _query_packages(self) -> None:
        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.all.packages.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn("Enviando query.all.packages.v1...")
        self._set_status("Buscando pacotes do SAP...")
        self._btn_packages.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_packages.configure(state=tk.NORMAL)

    def _query_versions_metadata(self) -> None:
        file_name = self._ver_filename_var.get().strip()
        category  = self._ver_category_var.get().strip()
        if not file_name:
            messagebox.showwarning("Atenção", "Informe o fileName para buscar versões.")
            return

        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "fileName": file_name,
                "category": category,
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.versions.metadata.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn(f"Enviando query.versions.metadata.v1 | fileName={file_name} | category={category}")
        self._set_status("Buscando versões do SAP...")
        self._btn_versions.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_versions.configure(state=tk.NORMAL)

    def _query_file_category(self) -> None:
        file_name = self._cat_filename_var.get().strip()
        if not file_name:
            messagebox.showwarning("Atenção", "Informe o fileName para buscar categoria.")
            return

        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "fileName": file_name,
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.file.category.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn(f"Enviando query.file.category.v1 | fileName={file_name}")
        self._set_status("Buscando categoria do SAP...")
        self._btn_category.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_category.configure(state=tk.NORMAL)

    def _query_request_files(self) -> None:
        raw = self._reqfiles_var.get().strip()
        if not raw:
            messagebox.showwarning("Atenção", "Informe pelo menos uma request.")
            return

        requests_list = [r.strip() for r in raw.split(",") if r.strip()]
        if not requests_list:
            messagebox.showwarning("Atenção", "Informe pelo menos uma request válida.")
            return

        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "requests": requests_list,
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.request.files.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn(f"Enviando query.request.files.v1 | requests={requests_list}")
        self._set_status("Buscando arquivos por request no SAP...")
        self._btn_request_files.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_request_files.configure(state=tk.NORMAL)

    def _query_request_description(self) -> None:
        request_id = self._reqdesc_var.get().strip()
        if not request_id:
            messagebox.showwarning("Atenção", "Informe o Request ID.")
            return

        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "requestId": request_id,
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.request.description.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn(f"Enviando query.request.description.v1 | requestId={request_id}")
        self._set_status("Buscando descrição da request no SAP...")
        self._btn_request_desc.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_request_desc.configure(state=tk.NORMAL)

    def _query_read_from_version(self) -> None:
        file_name  = self._rfv_filename_var.get().strip()
        category   = self._rfv_category_var.get().strip()
        version_id = self._rfv_version_var.get().strip()
        if not file_name:
            messagebox.showwarning("Atenção", "Informe o fileName.")
            return
        if not version_id:
            messagebox.showwarning("Atenção", "Informe o versionId.")
            return

        correlation_id = str(uuid.uuid4())
        v1_data = {
            "payload": {
                "fileName": file_name,
                "category": category,
                "versionId": version_id,
                "correlationId": correlation_id,
                "replyTo": "queue_vpn_respostas",
            },
            "routing_key": "usiminas.req.query.read.from.version.v1",
            "correlation_id": correlation_id,
        }
        self._log_fn(f"Enviando query.read.from.version.v1 | fileName={file_name} | category={category} | versionId={version_id}")
        self._set_status("Buscando conteúdo da versão no SAP...")
        self._btn_read_version.configure(state=tk.DISABLED)
        self._on_publish_v1(v1_data)
        self._btn_read_version.configure(state=tk.NORMAL)

    # ------------------------------------------------------------------ #
    #  Helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def _clear_panel(panel: tk.Text) -> None:
        panel.delete("1.0", tk.END)

    def _get_active_result_panel(self) -> tk.Text:
        """Retorna o painel de resultado da sub-aba ativa."""
        idx = self._notebook.index(self._notebook.select())
        return self._result_panels[idx]

    # ------------------------------------------------------------------ #
    #  Chamado pelo app.py quando a resposta chega
    # ------------------------------------------------------------------ #
    def display_response(self, response: list) -> None:
        """Exibe a resposta JSON formatada no painel da sub-aba ativa."""
        panel = self._get_active_result_panel()
        panel.delete("1.0", tk.END)
        formatted = json.dumps(response, indent=2, ensure_ascii=False)
        panel.insert("1.0", formatted)
        self._set_status(f"Query concluida: {len(response)} resultado(s)")

    def display_string_response(self, response: str) -> None:
        """Exibe resposta de string simples (ex: file category)."""
        panel = self._get_active_result_panel()
        panel.delete("1.0", tk.END)
        panel.insert("1.0", response)
        self._set_status(f"Query concluida: {response}")
