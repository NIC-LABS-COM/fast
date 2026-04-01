"""
Aba Queries — Testa operacoes de consulta (Requests e Reports).
Exibe o retorno JSON em um painel central.
"""
import json
import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict


class QueriesTab(ttk.Frame):
    def __init__(
        self,
        parent,
        on_publish_v1: Callable[[Dict], None],
        log_fn: Callable[[str], None],
        set_status_fn: Callable[[str], None],
        **kwargs,
    ) -> None:
        super().__init__(parent, padding=12, **kwargs)
        self._on_publish_v1 = on_publish_v1
        self._log_fn = log_fn
        self._set_status = set_status_fn
        self._build()

    def _build(self) -> None:
        # ---- Botoes no topo ----
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=(0, 8))

        self._btn_requests = ttk.Button(
            btn_frame, text="Buscar Requests", command=self._query_requests, width=22,
        )
        self._btn_requests.pack(side=tk.LEFT, padx=(0, 8))

        self._btn_reports = ttk.Button(
            btn_frame, text="Buscar Reports", command=self._query_reports, width=22,
        )
        self._btn_reports.pack(side=tk.LEFT, padx=(0, 8))

        self._btn_packages = ttk.Button(
            btn_frame, text="Buscar Pacotes", command=self._query_packages, width=22,
        )
        self._btn_packages.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(
            btn_frame, text="Limpar", command=self._clear, width=12,
        ).pack(side=tk.RIGHT)

        # ---- Versions Metadata ----
        versions_frame = ttk.LabelFrame(self, text="Versions Metadata")
        versions_frame.pack(fill=tk.X, pady=(0, 8))

        row = ttk.Frame(versions_frame)
        row.pack(fill=tk.X, padx=8, pady=6)

        ttk.Label(row, text="fileName:").pack(side=tk.LEFT)
        self._ver_filename_var = tk.StringVar()
        ttk.Entry(row, textvariable=self._ver_filename_var, width=30).pack(side=tk.LEFT, padx=(4, 12))

        ttk.Label(row, text="category:").pack(side=tk.LEFT)
        self._ver_category_var = tk.StringVar(value="PROGRAM")
        ttk.Combobox(
            row, textvariable=self._ver_category_var, width=20,
            values=["PROGRAM", "FUNCTION_MODULE", "CLASS"], state="readonly",
        ).pack(side=tk.LEFT, padx=(4, 12))

        self._btn_versions = ttk.Button(
            row, text="Buscar Versões", command=self._query_versions_metadata, width=18,
        )
        self._btn_versions.pack(side=tk.LEFT, padx=(4, 0))

        # ---- Painel de resultado no centro ----
        result_frame = ttk.LabelFrame(self, text="Retorno da Query")
        result_frame.pack(fill=tk.BOTH, expand=True)

        self._result_text = tk.Text(
            result_frame, wrap=tk.WORD, font=("Consolas", 10),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="#d4d4d4",
        )
        self._result_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8, side=tk.LEFT)

        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self._result_text.yview)
        scrollbar.pack(fill=tk.Y, side=tk.RIGHT, padx=(0, 8), pady=8)
        self._result_text.configure(yscrollcommand=scrollbar.set)

    # ------------------------------------------------------------------ #
    #  Acoes
    # ------------------------------------------------------------------ #
    def _query_requests(self) -> None:
        import uuid
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
        import uuid
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
        import uuid
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
            from tkinter import messagebox
            messagebox.showwarning("Atenção", "Informe o fileName para buscar versões.")
            return

        import uuid
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

    def _clear(self) -> None:
        self._result_text.delete("1.0", tk.END)

    # ------------------------------------------------------------------ #
    #  Chamado pelo app.py quando a resposta de query chega
    # ------------------------------------------------------------------ #
    def display_response(self, response: list) -> None:
        """Exibe a resposta JSON formatada no painel central."""
        self._result_text.delete("1.0", tk.END)
        formatted = json.dumps(response, indent=2, ensure_ascii=False)
        self._result_text.insert("1.0", formatted)
        self._set_status(f"Query concluida: {len(response)} resultado(s)")
