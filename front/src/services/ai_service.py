"""
Servico de integracao com a OpenAI API.
Gera propostas de tabelas SAP a partir de descricoes em linguagem natural.
"""
import json
import logging
import urllib.error
import urllib.request
from typing import Dict

from core.config import OPENAI_API_KEY, OPENAI_MODEL

logger = logging.getLogger("sap_publisher")

_SYSTEM_PROMPT_SE11 = (
    "Voce e um especialista em SAP ABAP Dictionary. Gere tabela transparente SAP.\n"
    "Retorne APENAS JSON valido (sem markdown):\n"
    '{"table_name":"Z...","table_text":"...","tabart":"APPL0","tabkat":"3",'
    '"fields":[{"field":"MANDT","key":"1","notnull":"1","element_data":"MANDT",'
    '"domain_data":"","desc":"Mandante","domtype":"","domlen":"","ref_table":"","ref_field":""}]}\n'
    "Regras: MANDT primeiro key=1 notnull=1; campos Z; field max 16; "
    "domtype CHAR/NUMC/DATS/DEC/INT4/CURR/QUAN; CURR/QUAN precisam ref; table_name Z max 16."
)

_SYSTEM_PROMPT_SE38 = (
    "Voce e um especialista em SAP ABAP. Gere programas ABAP para a transacao SE38.\n"
    "Retorne APENAS JSON valido (sem markdown) com esta estrutura exata:\n"
    '{"program_name":"Z...","title":"...","type":"1","status":"T","application":"",'
    '"auth_group":"","package":"$TMP","flow":"Criar e Ativar","source_code":"..."}\n'
    "Regras:\n"
    "- program_name: comeca com Z ou Y, max 40 chars, sem espacos\n"
    "- type: 1=Report, M=Module Pool, I=Include, S=Subroutine Pool, F=Function Group, K=Class Pool\n"
    "- status: T=Teste, P=Produtivo, S=Sistema\n"
    "- flow: CREATE / CREATE_ACTIVATE / Criar e Ativar\n"
    "- source_code: codigo ABAP completo e funcional, use \\n para quebras de linha\n"
    "- package: $TMP para objeto local, ou nome do pacote Z\n"
    "- Gere codigo ABAP real, funcional e bem comentado em portugues\n"
    "- Inclua REPORT statement, selecoes, logica e saida conforme solicitado"
)


class AIService:
    """Chama a API da OpenAI para gerar propostas SAP (tabelas e programas ABAP)."""

    def call(self, prompt: str, state: Dict) -> Dict:
        """Gera proposta de tabela SE11 (system prompt SE11)."""
        return self._request(_SYSTEM_PROMPT_SE11, prompt, state)

    def call_se38(self, prompt: str, state: Dict) -> Dict:
        """Gera proposta de programa ABAP SE38 (system prompt SE38)."""
        return self._request(_SYSTEM_PROMPT_SE38, prompt, state)

    def _request(self, system_prompt: str, prompt: str, state: Dict) -> Dict:
        """
        Envia prompt + estado atual da tabela para a OpenAI.
        Retorna dict com a proposta de tabela.
        Levanta excecao em caso de erro HTTP ou JSON invalido.
        """
        payload = {
            "model": OPENAI_MODEL,
            "messages": [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": (
                        f"Estado:\n{json.dumps(state, ensure_ascii=False, indent=2)}"
                        f"\n\nSolicitacao:\n{prompt}"
                    ),
                },
            ],
            "temperature": 0.2,
        }
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            "https://api.openai.com/v1/chat/completions",
            data=data,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {OPENAI_API_KEY}",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            rd = json.loads(resp.read().decode("utf-8"))

        content = rd["choices"][0]["message"]["content"].strip()
        content = self._strip_markdown(content)
        return json.loads(content)

    @staticmethod
    def _strip_markdown(content: str) -> str:
        """Remove blocos de codigo markdown que a API as vezes retorna."""
        if content.startswith("```"):
            parts = content.split("```")
            content = parts[1] if len(parts) > 1 else content
            if content.startswith("json"):
                content = content[4:]
        return content.strip()
