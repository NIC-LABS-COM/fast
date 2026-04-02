"""
Aba de criacao de Tabela Transparente SAP.
Inclui gerador assistido por IA e edicao manual de campos.
"""
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Dict, List

from core.config import DEFAULT_PACKAGE, DEFAULT_REQUEST, TABART_OPTIONS, TABKAT_OPTIONS, TYPES_REQUIRING_LENGTH
from services.ai_service import AIService
from services.publisher_service import PublisherService


class TabelaTab(ttk.Frame):
    def __init__(
        self,
        parent,
        publisher: PublisherService,
        ai: AIService,
        on_publish: Callable[[List[Dict]], None],
        log_fn: Callable[[str], None],
        set_status_fn: Callable[[str], None],
        **kwargs,
    ) -> None:
        super().__init__(parent, padding=12, **kwargs)
        self._publisher = publisher
        self._ai = ai
        self._on_publish = on_publish
        self._log_fn = log_fn
        self._set_status = set_status_fn
        self._build()

    # ------------------------------------------------------------------ #
    #  Build UI
    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        self._build_ai_frame()
        self._build_header_frame()
        self._build_grid_frame()

    def _build_ai_frame(self) -> None:
        ai_frame = ttk.LabelFrame(self, text="IA - Gerador de Tabela SAP")
        ai_frame.pack(fill=tk.X, pady=(0, 8))
        ai_frame.columnconfigure(0, weight=1)

        ttk.Label(ai_frame, text="Descreva a tabela que deseja criar:").grid(
            row=0, column=0, columnspan=2, padx=8, pady=(8, 2), sticky="w"
        )
        self.ai_prompt_text = tk.Text(ai_frame, height=3, wrap=tk.WORD, font=("TkDefaultFont", 9))
        self.ai_prompt_text.grid(row=1, column=0, padx=(8, 0), pady=(2, 8), sticky="ew")
        sc = ttk.Scrollbar(ai_frame, orient=tk.VERTICAL, command=self.ai_prompt_text.yview)
        self.ai_prompt_text.configure(yscrollcommand=sc.set)
        sc.grid(row=1, column=1, padx=(0, 4), pady=(2, 8), sticky="ns")

        btn_col = ttk.Frame(ai_frame)
        btn_col.grid(row=1, column=2, padx=8, pady=(2, 8), sticky="ns")
        self._btn_ai = ttk.Button(btn_col, text="Enviar para IA", command=self._send_to_ai, width=16)
        self._btn_ai.pack(fill=tk.X, pady=(0, 4))
        self._btn_tab = ttk.Button(btn_col, text="Enviar para Fila", command=self._publish, width=16)
        self._btn_tab.pack(fill=tk.X)

    def _build_header_frame(self) -> None:
        hd = ttk.LabelFrame(self, text="Cabecalho")
        hd.pack(fill=tk.X, pady=(0, 8))

        self.tab_name_var = tk.StringVar()
        self.tab_text_var = tk.StringVar()
        self.tab_package_var = tk.StringVar(value=DEFAULT_PACKAGE)
        self.tab_request_var = tk.StringVar(value=DEFAULT_REQUEST)
        self.tab_tabart_var = tk.StringVar(value="APPL0")
        self.tab_tabkat_var = tk.StringVar(value="3")

        ttk.Label(hd, text="Request:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.tab_request_var, width=20).grid(row=0, column=1, padx=8, pady=6, sticky="w")
        ttk.Label(hd, text="Pacote:").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.tab_package_var, width=20).grid(row=0, column=3, padx=8, pady=6, sticky="w")

        ttk.Label(hd, text="Tabela:").grid(row=1, column=0, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.tab_name_var, width=20).grid(row=1, column=1, padx=8, pady=6, sticky="w")
        ttk.Label(hd, text="Descricao:").grid(row=1, column=2, padx=8, pady=6, sticky="w")
        ttk.Entry(hd, textvariable=self.tab_text_var, width=40).grid(row=1, column=3, padx=8, pady=6, sticky="w")

        ttk.Label(hd, text="TABART:").grid(row=2, column=0, padx=8, pady=6, sticky="w")
        cb1 = ttk.Combobox(
            hd, textvariable=self.tab_tabart_var, state="readonly",
            width=12, values=list(TABART_OPTIONS.keys())
        )
        cb1.grid(row=2, column=1, padx=8, pady=6, sticky="w")
        cb1.bind("<<ComboboxSelected>>", lambda _: self._refresh_hints())

        ttk.Label(hd, text="TABKAT:").grid(row=2, column=2, padx=8, pady=6, sticky="w")
        cb2 = ttk.Combobox(
            hd, textvariable=self.tab_tabkat_var, state="readonly",
            width=12, values=list(TABKAT_OPTIONS.keys())
        )
        cb2.grid(row=2, column=3, padx=8, pady=6, sticky="w")
        cb2.bind("<<ComboboxSelected>>", lambda _: self._refresh_hints())

        self._hints_lbl = ttk.Label(hd, text="", anchor="w")
        self._hints_lbl.grid(row=3, column=0, columnspan=4, padx=8, pady=(2, 6), sticky="w")
        self._refresh_hints()

    def _build_grid_frame(self) -> None:
        gf = ttk.LabelFrame(self, text="Campos da Tabela")
        gf.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        cols = ("field", "key", "notnull", "element", "domain", "desc", "domtype", "domlen", "reftab", "reffield")
        self.tab_grid = ttk.Treeview(gf, columns=cols, show="headings", height=8)
        for c, t, w in [
            ("field", "FIELD", 160), ("key", "KEY", 50), ("notnull", "NOTNULL", 60),
            ("element", "ELEMENT", 150), ("domain", "DOMAIN", 150), ("desc", "DESC", 180),
            ("domtype", "TIPO", 60), ("domlen", "TAM", 50),
            ("reftab", "REF_TAB", 120), ("reffield", "REF_FIELD", 120),
        ]:
            self.tab_grid.heading(c, text=t)
            anc = "center" if c in ("key", "notnull", "domtype", "domlen") else "w"
            self.tab_grid.column(c, width=w, anchor=anc)

        ys = ttk.Scrollbar(gf, orient=tk.VERTICAL, command=self.tab_grid.yview)
        self.tab_grid.configure(yscrollcommand=ys.set)
        self.tab_grid.grid(row=0, column=0, columnspan=8, sticky="nsew", padx=8, pady=6)
        ys.grid(row=0, column=8, sticky="ns", pady=6)
        self.tab_grid.bind("<<TreeviewSelect>>", self._on_item_selected)
        gf.columnconfigure(0, weight=1)
        gf.rowconfigure(0, weight=1)

        # Variaveis de edicao de linha
        self.tf_field = tk.StringVar()
        self.tf_key = tk.BooleanVar(value=True)
        self.tf_notnull = tk.BooleanVar(value=True)
        self.tf_element = tk.StringVar()
        self.tf_domain = tk.StringVar()
        self.tf_desc = tk.StringVar()
        self.tf_domtype = tk.StringVar(value="CHAR")
        self.tf_domlen = tk.StringVar(value="15")
        self.tf_reftab = tk.StringVar()
        self.tf_reffield = tk.StringVar()

        ttk.Label(gf, text="FIELD:").grid(row=1, column=0, padx=4, pady=4, sticky="w")
        ttk.Entry(gf, textvariable=self.tf_field, width=16).grid(row=1, column=1, padx=4, pady=4, sticky="w")
        ttk.Checkbutton(gf, text="KEY", variable=self.tf_key).grid(row=1, column=2, padx=4, pady=4, sticky="w")
        ttk.Checkbutton(gf, text="NOTNULL", variable=self.tf_notnull).grid(row=1, column=3, padx=4, pady=4, sticky="w")
        ttk.Label(gf, text="ELEMENT:").grid(row=1, column=4, padx=4, pady=4, sticky="w")
        ttk.Entry(gf, textvariable=self.tf_element, width=16).grid(row=1, column=5, padx=4, pady=4, sticky="w")
        ttk.Label(gf, text="DOMAIN:").grid(row=1, column=6, padx=4, pady=4, sticky="w")
        ttk.Entry(gf, textvariable=self.tf_domain, width=16).grid(row=1, column=7, padx=4, pady=4, sticky="w")

        ttk.Label(gf, text="DESC:").grid(row=2, column=0, padx=4, pady=4, sticky="w")
        ttk.Entry(gf, textvariable=self.tf_desc, width=30).grid(row=2, column=1, columnspan=2, padx=4, pady=4, sticky="w")
        ttk.Label(gf, text="TIPO:").grid(row=2, column=3, padx=4, pady=4, sticky="w")
        ttk.Entry(gf, textvariable=self.tf_domtype, width=8).grid(row=2, column=4, padx=4, pady=4, sticky="w")
        ttk.Label(gf, text="TAM:").grid(row=2, column=5, padx=4, pady=4, sticky="w")
        ttk.Entry(gf, textvariable=self.tf_domlen, width=6).grid(row=2, column=6, padx=4, pady=4, sticky="w")

        br = ttk.Frame(gf)
        br.grid(row=3, column=0, columnspan=8, padx=4, pady=6, sticky="w")
        ttk.Button(br, text="Adicionar", command=self._add_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(br, text="Atualizar", command=self._update_row).pack(side=tk.LEFT, padx=4)
        ttk.Button(br, text="Remover", command=self._remove_row).pack(side=tk.LEFT, padx=4)

        # MANDT e obrigatorio em toda tabela transparente SAP
        self.tab_grid.insert("", tk.END, values=("MANDT", "1", "1", "MANDT", "", "Mandante", "", "", "", ""))

    # ------------------------------------------------------------------ #
    #  Header helpers
    # ------------------------------------------------------------------ #
    def _refresh_hints(self) -> None:
        a = TABART_OPTIONS.get(self.tab_tabart_var.get(), "")
        b = TABKAT_OPTIONS.get(self.tab_tabkat_var.get(), "")
        self._hints_lbl.configure(
            text=f"{self.tab_tabart_var.get()}: {a} | {self.tab_tabkat_var.get()}: {b}"
        )

    # ------------------------------------------------------------------ #
    #  Grid helpers
    # ------------------------------------------------------------------ #
    def _on_item_selected(self, _=None) -> None:
        sel = self.tab_grid.selection()
        if not sel:
            return
        v = self.tab_grid.item(sel[0], "values")
        self.tf_field.set(str(v[0]))
        self.tf_key.set(str(v[1]) == "1")
        self.tf_notnull.set(str(v[2]) == "1")
        self.tf_element.set(str(v[3]))
        self.tf_domain.set(str(v[4]) if len(v) > 4 else "")
        self.tf_desc.set(str(v[5]) if len(v) > 5 else "")
        self.tf_domtype.set(str(v[6]) if len(v) > 6 else "")
        self.tf_domlen.set(str(v[7]) if len(v) > 7 else "")
        self.tf_reftab.set(str(v[8]) if len(v) > 8 else "")
        self.tf_reffield.set(str(v[9]) if len(v) > 9 else "")

    def _row_values(self) -> tuple:
        el = self.tf_element.get().strip().upper()
        dom = self.tf_domain.get().strip().upper()
        if not dom and el.startswith("Z"):
            dom = el
        return (
            self.tf_field.get().strip().upper(),
            "1" if self.tf_key.get() else "0",
            "1" if self.tf_notnull.get() else "0",
            el, dom,
            self.tf_desc.get().strip(),
            self.tf_domtype.get().strip().upper(),
            self.tf_domlen.get().strip(),
            self.tf_reftab.get().strip().upper(),
            self.tf_reffield.get().strip().upper(),
        )

    def _add_row(self) -> None:
        v = self._row_values()
        if not v[0] or not v[3]:
            return messagebox.showwarning("Validacao", "Informe FIELD e ELEMENT.")
        domtype = str(v[6]).strip().upper()
        domlen = str(v[7]).strip()
        if domtype in TYPES_REQUIRING_LENGTH and (not domlen.isdigit() or int(domlen) < 1):
            return messagebox.showwarning(
                "Validacao",
                f"Tipo {domtype} exige tamanho numerico maior que zero.",
            )
        self.tab_grid.insert("", tk.END, values=v)

    def _update_row(self) -> None:
        sel = self.tab_grid.selection()
        if not sel:
            return messagebox.showwarning("Atencao", "Selecione um item.")
        v = self._row_values()
        if not v[0] or not v[3]:
            return messagebox.showwarning("Validacao", "Informe FIELD e ELEMENT.")
        self.tab_grid.item(sel[0], values=v)

    def _remove_row(self) -> None:
        sel = self.tab_grid.selection()
        if not sel:
            return messagebox.showwarning("Atencao", "Selecione um item.")
        self.tab_grid.delete(sel[0])

    def _collect_items(self) -> List[Dict]:
        rows = []
        for rid in self.tab_grid.get_children():
            v = self.tab_grid.item(rid, "values")
            if len(v) < 4 or not str(v[0]).strip() or not str(v[3]).strip():
                continue
            rows.append({
                "field": str(v[0]).strip().upper(),
                "key": str(v[1]).strip(),
                "notnull": str(v[2]).strip(),
                "element_data": str(v[3]).strip().upper(),
                "domain_data": str(v[4]).strip().upper() if len(v) > 4 else "",
                "desc": str(v[5]).strip() if len(v) > 5 else "",
                "domtype": str(v[6]).strip().upper() if len(v) > 6 else "",
                "domlen": str(v[7]).strip() if len(v) > 7 else "",
                "ref_table": str(v[8]).strip().upper() if len(v) > 8 else "",
                "ref_field": str(v[9]).strip().upper() if len(v) > 9 else "",
            })
        return rows

    def _collect_ai_state(self) -> Dict:
        return {
            "table_name": self.tab_name_var.get().strip(),
            "table_text": self.tab_text_var.get().strip(),
            "tabart": self.tab_tabart_var.get().strip(),
            "tabkat": self.tab_tabkat_var.get().strip(),
            "fields": self._collect_items(),
        }

    # ------------------------------------------------------------------ #
    #  Publicar tabela
    # ------------------------------------------------------------------ #
    def _publish(self) -> None:
        n = self.tab_name_var.get().strip()
        t = self.tab_text_var.get().strip()
        pk = self.tab_package_var.get().strip()
        rq = self.tab_request_var.get().strip()

        if not n:
            return messagebox.showwarning("Validacao", "Informe o nome da tabela.")
        if not t:
            return messagebox.showwarning("Validacao", "Informe a descricao.")
        if not pk:
            return messagebox.showwarning("Validacao", "Informe o pacote.")
        if not rq:
            return messagebox.showwarning("Validacao", "Informe a request.")

        items = self._collect_items()
        if not items:
            return messagebox.showwarning("Validacao", "Adicione ao menos um campo.")
        if len(items) > 12:
            return messagebox.showwarning("Validacao", "Limite de 12 campos.")

        messages = self._publisher.build_message_chain(
            items, n.upper(), t,
            self.tab_tabart_var.get(), self.tab_tabkat_var.get(),
            pk, rq.upper(),
        )
        n_dom = sum(1 for m in messages if m["action"] == "criar_dominio")
        n_elem = sum(1 for m in messages if m["action"] == "criar_elemento")
        resumo = (
            f"Serao enviadas {len(messages)} mensagens:\n\n"
            f"  - {n_dom} dominio(s)\n  - {n_elem} elemento(s)\n  - 1 tabela ({n.upper()})\n\n"
            f"Confirma o envio?"
        )
        if not messagebox.askyesno("Confirmar", resumo):
            return

        self._btn_tab.configure(state=tk.DISABLED)
        self._log_fn(f"Enviando cadeia: {n_dom} dominios + {n_elem} elementos + 1 tabela...")
        try:
            self._on_publish(messages)
        finally:
            self._btn_tab.configure(state=tk.NORMAL)

    # ------------------------------------------------------------------ #
    #  IA
    # ------------------------------------------------------------------ #
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
            result = self._ai.call(prompt, state)
            self.after(0, self._apply_ai, result)
        except Exception as exc:
            self.after(0, self._ai_error, str(exc))

    def _apply_ai(self, result: Dict) -> None:
        try:
            if "table_name" in result:
                self.tab_name_var.set(str(result["table_name"]).upper()[:16])
            if "table_text" in result:
                self.tab_text_var.set(str(result["table_text"])[:60])
            if "tabart" in result and str(result["tabart"]) in TABART_OPTIONS:
                self.tab_tabart_var.set(str(result["tabart"]))
            if "tabkat" in result and str(result["tabkat"]) in TABKAT_OPTIONS:
                self.tab_tabkat_var.set(str(result["tabkat"]))
            self._refresh_hints()

            fields = result.get("fields", [])
            if fields:
                for rid in self.tab_grid.get_children():
                    self.tab_grid.delete(rid)
                for f in fields:
                    self.tab_grid.insert("", tk.END, values=(
                        str(f.get("field", "")).upper()[:16],
                        "1" if str(f.get("key", "0")) == "1" else "0",
                        "1" if str(f.get("notnull", "0")) == "1" else "0",
                        str(f.get("element_data", "")).upper()[:30],
                        str(f.get("domain_data", "")).upper()[:30],
                        str(f.get("desc", ""))[:60],
                        str(f.get("domtype", "")).upper()[:10],
                        str(f.get("domlen", ""))[:5],
                        str(f.get("ref_table", "")).upper()[:30],
                        str(f.get("ref_field", "")).upper()[:30],
                    ))

            self._set_status("IA respondeu. Revise antes de enviar.")
            messagebox.showinfo("IA", "Proposta gerada! Revise antes de enviar.")
        except Exception as exc:
            self._ai_error(f"Erro ao aplicar resultado: {exc}")
        finally:
            self._btn_ai.configure(state=tk.NORMAL)

    def _ai_error(self, msg: str) -> None:
        self._set_status("Erro IA")
        messagebox.showerror("Erro IA", msg)
        self._btn_ai.configure(state=tk.NORMAL)
