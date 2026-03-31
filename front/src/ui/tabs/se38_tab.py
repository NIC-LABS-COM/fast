"""
Aba SE38 - Criacao de Programas ABAP.
Somente UI e validacao; logica delegada ao PublisherService e AIService.
"""
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Dict, List

from core.config import (
    DEFAULT_PACKAGE, DEFAULT_REQUEST,
    PROG_FLOW_OPTIONS, PROG_TYPE_OPTIONS,
)
from services.ai_service import AIService
from services.publisher_service import PublisherService
from ui.widgets.comparison_panel import ComparisonPanel


class SE38Tab(ttk.Frame):
    def __init__(
        self,
        parent,
        publisher: PublisherService,
        ai: AIService,
        on_publish: Callable[[List[Dict]], None],
        on_publish_v1: Callable[[Dict], None],
        log_fn: Callable[[str], None],
        set_status_fn: Callable[[str], None],
        **kwargs,
    ) -> None:
        super().__init__(parent, padding=12, **kwargs)
        self._publisher     = publisher
        self._ai            = ai
        self._on_publish    = on_publish
        self._on_publish_v1 = on_publish_v1
        self._log_fn        = log_fn
        self._set_status    = set_status_fn
        self._build()

    # ------------------------------------------------------------------ #
    #  Build UI
    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        self._build_ai_frame()
        self._build_header()
        self._build_code_frame()
        self._build_actions()

    def _build_ai_frame(self) -> None:
        ai_frame = ttk.LabelFrame(self, text="IA - Gerador de Programa ABAP")
        ai_frame.pack(fill=tk.X, pady=(0, 8))
        ai_frame.columnconfigure(0, weight=1)

        ttk.Label(ai_frame, text="Descreva o programa ABAP que deseja criar:").grid(
            row=0, column=0, columnspan=2, padx=8, pady=(8, 2), sticky="w"
        )
        self.ai_prompt_text = tk.Text(ai_frame, height=3, wrap=tk.WORD, font=("TkDefaultFont", 9))
        self.ai_prompt_text.grid(row=1, column=0, padx=(8, 0), pady=(2, 8), sticky="ew")
        sc = ttk.Scrollbar(ai_frame, orient=tk.VERTICAL, command=self.ai_prompt_text.yview)
        self.ai_prompt_text.configure(yscrollcommand=sc.set)
        sc.grid(row=1, column=1, padx=(0, 4), pady=(2, 8), sticky="ns")

        self._btn_ai = ttk.Button(ai_frame, text="Enviar para IA", command=self._send_to_ai, width=16)
        self._btn_ai.grid(row=1, column=2, padx=8, pady=(2, 8), sticky="n")

    def _build_header(self) -> None:
        hd = ttk.LabelFrame(self, text="Parametros do Programa")
        hd.pack(fill=tk.X, pady=(0, 8))

        self.prog_name_var    = tk.StringVar()
        self.prog_title_var   = tk.StringVar()
        self.prog_type_var    = tk.StringVar(value=PROG_TYPE_OPTIONS["1"])
        self.prog_appl_var    = tk.StringVar()
        self.prog_auth_var    = tk.StringVar()
        self.prog_package_var = tk.StringVar(value=DEFAULT_PACKAGE)
        self.prog_request_var = tk.StringVar(value=DEFAULT_REQUEST)
        self.prog_flow_var    = tk.StringVar(value="Criar e Ativar")

        # Linha 0 — Nome e Titulo
        ttk.Label(hd, text="Nome do programa:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.prog_name_var, width=22).grid(row=0, column=1, padx=8, pady=6, sticky="w")
        self._btn_buscar = ttk.Button(hd, text="Buscar Programa", command=self._buscar_no_sap, width=16)
        self._btn_buscar.grid(row=0, column=2, padx=(0, 8), pady=6, sticky="w")
        self._btn_buscar.grid_remove()  # visivel apenas no fluxo FIX
        ttk.Label(hd, text="Titulo:").grid(row=0, column=3, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.prog_title_var, width=36).grid(row=0, column=4, columnspan=2, padx=8, pady=6, sticky="w")

        # Linha exclusiva FIX — removida (não é necessária)
        # Mantém apenas o botão "Buscar Programa" que já existe

        # Linha 1 — Tipo e Status
        ttk.Label(hd, text="Tipo:").grid(row=1, column=0, padx=8, pady=6, sticky="w")
        cb_type = ttk.Combobox(
            hd, textvariable=self.prog_type_var, state="readonly", width=32,
            values=list(PROG_TYPE_OPTIONS.values()),
        )
        cb_type.grid(row=1, column=1, columnspan=2, padx=8, pady=6, sticky="w")

        # Linha 2 — Aplicacao e Grupo Auth
        ttk.Label(hd, text="Aplicacao:").grid(row=2, column=0, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.prog_appl_var, width=10).grid(row=2, column=1, padx=8, pady=6, sticky="w")
        ttk.Label(hd, text="Grupo Auth:").grid(row=2, column=2, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.prog_auth_var, width=10).grid(row=2, column=3, padx=8, pady=6, sticky="w")

        # Linha 3 — Pacote e Request
        ttk.Label(hd, text="Pacote:").grid(row=3, column=0, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.prog_package_var, width=22).grid(row=3, column=1, padx=8, pady=6, sticky="w")
        ttk.Label(hd, text="Request:").grid(row=3, column=2, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.prog_request_var, width=22).grid(row=3, column=3, padx=8, pady=6, sticky="w")

        # Linha 4 — Fluxo
        ttk.Label(hd, text="Fluxo:").grid(row=4, column=0, padx=8, pady=6, sticky="w")
        cb_flow = ttk.Combobox(
            hd, textvariable=self.prog_flow_var, state="readonly", width=35,
            values=list(PROG_FLOW_OPTIONS.keys()),
        )
        cb_flow.grid(row=4, column=1, columnspan=2, padx=8, pady=6, sticky="w")
        self._flow_hint = ttk.Label(hd, text=PROG_FLOW_OPTIONS["Criar e Ativar"], foreground="#555")
        self._flow_hint.grid(row=4, column=3, columnspan=3, padx=8, sticky="w")
        cb_flow.bind("<<ComboboxSelected>>", self._on_flow_change)

    def _build_code_frame(self) -> None:
        cf = ttk.LabelFrame(self, text="Codigo Fonte ABAP")
        cf.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        cf.columnconfigure(0, weight=1)
        cf.rowconfigure(0, weight=1)

        # Substituir Text simples por ComparisonPanel
        self.code_panel = ComparisonPanel(cf)
        self.code_panel.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)
        
        # Compatibilidade: criar alias code_text -> code_panel
        self.code_text = self.code_panel

        btn_row = ttk.Frame(cf)
        btn_row.grid(row=1, column=0, sticky="w", padx=8, pady=(0, 6))
        ttk.Button(btn_row, text="Carregar de Arquivo", command=self._load_file).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="Salvar em Arquivo",   command=self._save_file).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="Limpar Codigo",       command=self._clear_code).pack(side=tk.LEFT, padx=4)

    def _build_actions(self) -> None:
        self._btn = ttk.Button(self, text="Enviar para Fila", command=self._publish)
        self._btn.pack(anchor="w", pady=(0, 4))

    # ------------------------------------------------------------------ #
    #  Eventos
    # ------------------------------------------------------------------ #
    def _on_flow_change(self, _=None) -> None:
        flow = self.prog_flow_var.get()
        self._flow_hint.configure(text=PROG_FLOW_OPTIONS.get(flow, ""))
        if flow == "FIX":
            self._btn_buscar.grid()
        else:
            self._btn_buscar.grid_remove()

    def _load_file(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("ABAP/Texto", "*.abap *.txt"), ("Todos", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            self.code_text.delete("1.0", tk.END)
            self.code_text.insert("1.0", content)
            self._log_fn(f"Codigo carregado de: {path}")
        except Exception as exc:
            messagebox.showerror("Erro", f"Nao foi possivel ler o arquivo:\n{exc}")

    def _save_file(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".abap",
            filetypes=[("ABAP", "*.abap"), ("Texto", "*.txt"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.code_text.get("1.0", tk.END))
            self._log_fn(f"Codigo salvo em: {path}")
        except Exception as exc:
            messagebox.showerror("Erro", f"Nao foi possivel salvar:\n{exc}")

    def _clear_code(self) -> None:
        self.code_panel.clear_all()

    # ------------------------------------------------------------------ #
    #  IA
    # ------------------------------------------------------------------ #
    def _collect_ai_state(self) -> Dict:
        return {
            "program_name": self.prog_name_var.get().strip(),
            "title":        self.prog_title_var.get().strip(),
            "type":         self.prog_type_var.get().strip(),
            "application":  self.prog_appl_var.get().strip(),
            "auth_group":   self.prog_auth_var.get().strip(),
            "package":      self.prog_package_var.get().strip(),
            "flow":         self.prog_flow_var.get().strip(),
            "source_code":  self.code_text.get("1.0", tk.END).strip(),
        }

    def _send_to_ai(self) -> None:
        prompt = self.ai_prompt_text.get("1.0", tk.END).strip()
        if not prompt:
            return messagebox.showwarning("IA", "Informe um prompt.")
        self._btn_ai.configure(state=tk.DISABLED)
        self._set_status("Aguardando IA...")
        threading.Thread(
            target=self._call_ai_thread,
            args=(prompt, self._collect_ai_state()),
            daemon=True,
        ).start()

    def _call_ai_thread(self, prompt: str, state: Dict) -> None:
        try:
            result = self._ai.call_se38(prompt, state)
            self.after(0, self._apply_ai, result)
        except Exception as exc:
            self.after(0, self._ai_error, str(exc))

    def _apply_ai(self, result: Dict) -> None:
        try:
            # Preserva o flow atual se houver código original (modo FIX ativo)
            preserve_flow = self.code_panel.get_original() is not None
            
            if "program_name" in result:
                self.prog_name_var.set(str(result["program_name"]).upper()[:40])
            if "title" in result:
                self.prog_title_var.set(str(result["title"])[:60])
            if "type" in result and str(result["type"]) in PROG_TYPE_OPTIONS:
                self.prog_type_var.set(PROG_TYPE_OPTIONS[str(result["type"])])
            if "application" in result:
                self.prog_appl_var.set(str(result["application"]))
            if "auth_group" in result:
                self.prog_auth_var.set(str(result["auth_group"]))
            if "package" in result:
                self.prog_package_var.set(str(result["package"]))
            
            # Só atualiza flow se não estiver preservando (não estiver no modo FIX)
            if not preserve_flow and "flow" in result and str(result["flow"]) in PROG_FLOW_OPTIONS:
                self.prog_flow_var.set(str(result["flow"]))
                self._on_flow_change()
            
            if "source_code" in result and result["source_code"]:
                code = str(result["source_code"]).replace("\\n", "\n")
                
                # Se estiver no modo FIX e houver codigo original, mostra comparacao
                flow = self.prog_flow_var.get().strip()
                original_code = self.code_panel.get_original()
                
                # DEBUG: Log para verificar condições
                self._log_fn(f"DEBUG: flow='{flow}', tem_original={bool(original_code)}")
                
                if flow == "FIX" and original_code:
                    # DIVIDE A TELA: codigo da IA no painel MODIFICADO (direito)
                    self.code_panel.set_modified(code)
                    self._log_fn("✓ Tela dividida! Compare: Original (esquerda) vs IA (direita).")
                else:
                    # Modo normal: substitui codigo unico
                    self._log_fn(f"DEBUG: Entrando no modo normal (não dividiu)")
                    self.code_text.delete("1.0", tk.END)
                    self.code_text.insert("1.0", code)

            self._set_status("IA respondeu. Revise antes de enviar.")
            messagebox.showinfo("IA", "Proposta gerada! Revise o codigo e os campos antes de enviar.")
        except Exception as exc:
            self._ai_error(f"Erro ao aplicar resultado: {exc}")
        finally:
            self._btn_ai.configure(state=tk.NORMAL)

    def _ai_error(self, msg: str) -> None:
        self._set_status("Erro IA")
        messagebox.showerror("Erro IA", msg)
        self._btn_ai.configure(state=tk.NORMAL)

    # ------------------------------------------------------------------ #
    #  FIX — buscar codigo do SAP
    # ------------------------------------------------------------------ #
    def _buscar_no_sap(self) -> None:
        name = self.prog_name_var.get().strip().upper()
        if not name:
            return messagebox.showwarning("Validacao", "Informe o nome do programa antes de buscar.")
        if not name.startswith(("Z", "Y")):
            return messagebox.showwarning("Validacao", "Nome deve comecar com Z ou Y.")
        
        # Força fluxo FIX quando busca programa do SAP
        self.prog_flow_var.set("FIX")
        self._on_flow_change()
        
        v1_data = self._publisher.build_read_file_v1(name)
        self._log_fn(f"Buscando codigo de {name} no SAP...")
        self._btn_buscar.configure(state=tk.DISABLED)
        try:
            # Envia via V1 usando exchange topic
            self._on_publish_v1(v1_data)
        finally:
            self._btn_buscar.configure(state=tk.NORMAL)

    def load_source_code(self, code: str) -> None:
        """Chamado pelo app.py quando a resposta de buscar_programa chega."""
        # Carrega código no painel único (ainda não divide a tela)
        self.code_panel.set_original(code)
        # Desativa modo comparação até a IA gerar código
        self.code_panel.disable_comparison_mode()
        self._log_fn("Codigo carregado do SAP. Use 'Enviar para IA' para gerar versao modificada.")

    # ------------------------------------------------------------------ #
    #  Publicacao
    # ------------------------------------------------------------------ #
    def _publish(self) -> None:
        name    = self.prog_name_var.get().strip().upper()
        flow    = self.prog_flow_var.get().strip()
        code    = self.code_text.get("1.0", tk.END).strip()
        package = self.prog_package_var.get().strip()
        request = self.prog_request_var.get().strip().upper()

        if not name:
            return messagebox.showwarning("Validacao", "Informe o nome do programa.")
        if not name.startswith(("Z", "Y")):
            return messagebox.showwarning("Validacao", "Nome deve comecar com Z ou Y.")
        if len(name) > 40:
            return messagebox.showwarning("Validacao", "Nome do programa muito longo (max 40).")

        if flow == "FIX":
            code = self.code_text.get("1.0", tk.END).strip()
            if not code:
                return messagebox.showwarning("Validacao", "Carregue o codigo antes de enviar (use Buscar Programa).")
            package = self.prog_package_var.get().strip()
            request = self.prog_request_var.get().strip().upper()
            v1_data = self._publisher.build_fix_v1_payload(name, code)
            self._log_fn(f"Enviando edicao de {name} via V1...")
            self._btn.configure(state=tk.DISABLED)
            try:
                self._on_publish_v1(v1_data)
            finally:
                self._btn.configure(state=tk.NORMAL)
            return

        # ---- Fluxos de criacao ----
        title   = self.prog_title_var.get().strip()
        ptype   = next((k for k, v in PROG_TYPE_OPTIONS.items() if v == self.prog_type_var.get()), "1")
        pstatus = ""
        appl    = self.prog_appl_var.get().strip()
        auth    = self.prog_auth_var.get().strip()

        if not title:
            return messagebox.showwarning("Validacao", "Informe o titulo do programa.")
        if not package:
            return messagebox.showwarning("Validacao", "Informe o pacote ($TMP para local).")
        if package != "$TMP" and not request:
            return messagebox.showwarning("Validacao", "Request obrigatoria quando pacote nao e $TMP.")
        if "CODE" in flow and not code:
            return messagebox.showwarning("Validacao", "Fluxo selecionado exige codigo fonte.")

        payload = self._publisher.build_programa_payload(
            name, title, ptype, pstatus, appl, auth, package, request, flow, code
        )

        self._log_fn(f"Enviando programa {name} para fila (fluxo: {PROG_FLOW_OPTIONS[flow]})...")
        self._btn.configure(state=tk.DISABLED)
        try:
            self._on_publish([payload])
        finally:
            self._btn.configure(state=tk.NORMAL)
