"""
Aba de criacao de Elemento de Dados SAP.
Somente UI e validacao; logica de negocio delegada ao PublisherService.
"""
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Dict, List

from core.config import DEFAULT_PACKAGE, DEFAULT_REQUEST
from services.publisher_service import PublisherService


class ElementoTab(ttk.Frame):
    def __init__(
        self,
        parent,
        publisher: PublisherService,
        on_publish: Callable[[List[Dict]], None],
        log_fn: Callable[[str], None],
        **kwargs,
    ) -> None:
        super().__init__(parent, padding=12, **kwargs)
        self._publisher = publisher
        self._on_publish = on_publish
        self._log_fn = log_fn
        self._build()

    def _build(self) -> None:
        form = ttk.LabelFrame(self, text="Parametros do Elemento de Dados")
        form.pack(fill=tk.X, pady=(0, 10))

        self.elem_name_var = tk.StringVar()
        self.elem_text_var = tk.StringVar()
        self.elem_domain_var = tk.StringVar()
        self.elem_package_var = tk.StringVar(value=DEFAULT_PACKAGE)
        self.elem_request_var = tk.StringVar(value=DEFAULT_REQUEST)

        ttk.Label(form, text="Nome do elemento:").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(form, textvariable=self.elem_name_var, width=28).grid(row=0, column=1, padx=8, pady=8, sticky="w")
        ttk.Label(form, text="Descricao:").grid(row=0, column=2, padx=8, pady=8, sticky="w")
        ttk.Entry(form, textvariable=self.elem_text_var, width=40).grid(row=0, column=3, padx=8, pady=8, sticky="w")

        ttk.Label(form, text="Dominio:").grid(row=1, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(form, textvariable=self.elem_domain_var, width=28).grid(row=1, column=1, padx=8, pady=8, sticky="w")

        ttk.Label(form, text="Pacote:").grid(row=2, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(form, textvariable=self.elem_package_var, width=28).grid(row=2, column=1, padx=8, pady=8, sticky="w")
        ttk.Label(form, text="Request:").grid(row=2, column=2, padx=8, pady=8, sticky="w")
        ttk.Entry(form, textvariable=self.elem_request_var, width=28).grid(row=2, column=3, padx=8, pady=8, sticky="w")

        self._btn = ttk.Button(self, text="Enviar para Fila", command=self._publish)
        self._btn.pack(anchor="w", pady=8)

        info = ttk.LabelFrame(self, text="Observacoes")
        info.pack(fill=tk.BOTH, expand=True)
        ttk.Label(info, text=(
            "1. O elemento sera criado apontando para o dominio informado.\n"
            "2. Certifique-se de que o dominio ja existe antes de criar o elemento.\n"
            "3. Acompanhe o resultado no Log de Respostas abaixo."
        ), justify=tk.LEFT, anchor="w").pack(fill=tk.BOTH, padx=10, pady=10)

    def _publish(self) -> None:
        name = self.elem_name_var.get().strip().upper()
        text = self.elem_text_var.get().strip()
        domain = self.elem_domain_var.get().strip().upper()
        package = self.elem_package_var.get().strip()
        request = self.elem_request_var.get().strip().upper()

        if not name:
            return messagebox.showwarning("Validacao", "Informe o nome.")
        if not text:
            return messagebox.showwarning("Validacao", "Informe a descricao.")
        if not domain:
            return messagebox.showwarning("Validacao", "Informe o dominio.")
        if not package:
            return messagebox.showwarning("Validacao", "Informe o pacote.")
        if not request:
            return messagebox.showwarning("Validacao", "Informe a request.")

        payload = self._publisher.build_elemento_payload(name, text, domain, package, request)
        self._log_fn(f"Enviando elemento {name} para fila...")

        self._btn.configure(state=tk.DISABLED)
        try:
            self._on_publish([payload])
        finally:
            self._btn.configure(state=tk.NORMAL)
