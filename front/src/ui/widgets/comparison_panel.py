"""
ComparisonPanel - Widget para comparacao lado a lado de codigo.
Permite visualizar codigo ORIGINAL vs codigo MODIFICADO (IA).
"""
import tkinter as tk
from tkinter import ttk
from typing import Optional


class ComparisonPanel(ttk.Frame):
    """Painel de comparacao com dois editores de codigo lado a lado."""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._comparison_mode = False  # False = painel unico, True = comparacao
        self._build()
    
    def _build(self) -> None:
        """Constroi a interface do painel."""
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        # Container principal que alterna entre modo unico e comparacao
        self._single_container = ttk.Frame(self)
        self._comparison_container = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        
        # === MODO UNICO - Editor principal ===
        self._build_single_editor()
        
        # === MODO COMPARACAO - Dois editores lado a lado ===
        self._build_comparison_editors()
        
        # Inicia em modo unico
        self._show_single_mode()
    
    def _build_single_editor(self) -> None:
        """Constroi o editor unico (modo padrao)."""
        frame = ttk.Frame(self._single_container)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        self._single_text = tk.Text(
            frame, wrap=tk.NONE, font=("Courier New", 9),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
        )
        self._single_text.grid(row=0, column=0, sticky="nsew")
        
        ys = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self._single_text.yview)
        xs = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self._single_text.xview)
        self._single_text.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
    
    def _build_comparison_editors(self) -> None:
        """Constroi os dois editores para modo comparacao."""
        # === PAINEL ESQUERDO - Codigo Original ===
        left_frame = ttk.LabelFrame(self._comparison_container, text="Código Original")
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        self._original_text = tk.Text(
            left_frame, wrap=tk.NONE, font=("Courier New", 9),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
        )
        self._original_text.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
        
        ys_left = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self._original_text.yview)
        xs_left = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, command=self._original_text.xview)
        self._original_text.configure(yscrollcommand=ys_left.set, xscrollcommand=xs_left.set)
        ys_left.grid(row=0, column=1, sticky="ns", pady=4)
        xs_left.grid(row=1, column=0, sticky="ew", padx=(4, 0))
        
        # === PAINEL DIREITO - Codigo Modificado (IA) ===
        right_frame = ttk.LabelFrame(self._comparison_container, text="Código Modificado (IA)")
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        
        self._modified_text = tk.Text(
            right_frame, wrap=tk.NONE, font=("Courier New", 9),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
        )
        self._modified_text.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
        
        ys_right = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self._modified_text.yview)
        xs_right = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self._modified_text.xview)
        self._modified_text.configure(yscrollcommand=ys_right.set, xscrollcommand=xs_right.set)
        ys_right.grid(row=0, column=1, sticky="ns", pady=4)
        xs_right.grid(row=1, column=0, sticky="ew", padx=(4, 0))
        
        # Adiciona os paineis ao PanedWindow
        self._comparison_container.add(left_frame, weight=1)
        self._comparison_container.add(right_frame, weight=1)
    
    def _show_single_mode(self) -> None:
        """Mostra apenas o editor unico."""
        self._comparison_container.grid_forget()
        self._single_container.grid(row=0, column=0, sticky="nsew")
        self._comparison_mode = False
    
    def _show_comparison_mode(self) -> None:
        """Mostra os dois editores lado a lado."""
        self._single_container.grid_forget()
        self._comparison_container.grid(row=0, column=0, sticky="nsew")
        self._comparison_mode = True
    
    # ================================================================== #
    #  API PUBLICA - Compatibilidade com Text widget
    # ================================================================== #
    
    def get(self, start: str, end: Optional[str] = None) -> str:
        """Retorna texto do editor ativo (compativel com tk.Text)."""
        if self._comparison_mode:
            # No modo comparacao, retorna o codigo MODIFICADO (que sera enviado)
            return self._modified_text.get(start, end) if end else self._modified_text.get(start)
        return self._single_text.get(start, end) if end else self._single_text.get(start)
    
    def delete(self, start: str, end: Optional[str] = None) -> None:
        """Deleta texto do editor ativo (compativel com tk.Text)."""
        if self._comparison_mode:
            self._modified_text.delete(start, end)
        else:
            self._single_text.delete(start, end)
    
    def insert(self, index: str, text: str) -> None:
        """Insere texto no editor ativo (compativel com tk.Text)."""
        if self._comparison_mode:
            self._modified_text.insert(index, text)
        else:
            self._single_text.insert(index, text)
    
    # ================================================================== #
    #  API PUBLICA - Controle de Comparacao
    # ================================================================== #
    
    def set_original(self, code: str) -> None:
        """Define o codigo original (painel esquerdo)."""
        self._original_text.delete("1.0", tk.END)
        self._original_text.insert("1.0", code)
        # Copia para o editor unico tambem
        self._single_text.delete("1.0", tk.END)
        self._single_text.insert("1.0", code)
    
    def set_modified(self, code: str) -> None:
        """Define o codigo modificado (painel direito) e ativa modo comparacao."""
        self._modified_text.delete("1.0", tk.END)
        self._modified_text.insert("1.0", code)
        self.enable_comparison_mode()
    
    def get_original(self) -> str:
        """Retorna o codigo original."""
        return self._original_text.get("1.0", tk.END).strip()
    
    def get_modified(self) -> str:
        """Retorna o codigo modificado."""
        return self._modified_text.get("1.0", tk.END).strip()
    
    def enable_comparison_mode(self) -> None:
        """Ativa o modo comparacao (dois paineis)."""
        self._show_comparison_mode()
    
    def disable_comparison_mode(self) -> None:
        """Desativa o modo comparacao (volta para painel unico)."""
        self._show_single_mode()
    
    def clear_all(self) -> None:
        """Limpa todos os editores e volta para modo unico."""
        self._single_text.delete("1.0", tk.END)
        self._original_text.delete("1.0", tk.END)
        self._modified_text.delete("1.0", tk.END)
        self.disable_comparison_mode()
    
    def is_comparison_mode(self) -> bool:
        """Retorna True se estiver em modo comparacao."""
        return self._comparison_mode
