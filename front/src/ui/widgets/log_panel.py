"""
Widget reutilizavel de log de respostas do SAP.
"""
import tkinter as tk
from datetime import datetime
from tkinter import ttk


class LogPanel(ttk.LabelFrame):
    """Painel de log colorido exibido na parte inferior da janela principal."""

    def __init__(self, parent, **kwargs) -> None:
        super().__init__(parent, text="Log de Respostas do SAP", **kwargs)
        self._build()

    def _build(self) -> None:
        self.log_text = tk.Text(
            self, height=6, wrap=tk.WORD, font=("Consolas", 9), state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.X, padx=8, pady=6, side=tk.LEFT, expand=True)

        scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=6)

        self.log_text.tag_configure("sucesso", foreground="#2e7d32")
        self.log_text.tag_configure("ja_existe", foreground="#f57f17")
        self.log_text.tag_configure("erro", foreground="#c62828")

    def append(self, text: str, status: str = "") -> None:
        """Adiciona uma linha ao log, opcionalmente colorida pelo status."""
        self.log_text.configure(state=tk.NORMAL)
        tag = status if status in ("sucesso", "ja_existe", "erro") else ""
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {text}\n"
        if tag:
            self.log_text.insert(tk.END, line, tag)
        else:
            self.log_text.insert(tk.END, line)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def clear(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)
