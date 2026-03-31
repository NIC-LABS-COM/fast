"""
Servico de negocio para construcao de payloads SAP.
Monta cadeias de mensagens (dominio -> elemento -> tabela).
"""
from typing import Dict, List, Set

from core.config import DEFAULT_DOMAIN_DATATYPE, DEFAULT_DOMAIN_LENGTH, VBS_URLS, PROG_FLOW_OPTIONS, QUEUE_RESPONSES
from models.messages import SapCommand, SapEventV1


class PublisherService:
    """Constroe os payloads enviados para o consumer SAP via RabbitMQ."""

    def build_dominio_payload(
        self,
        name: str,
        text: str,
        datatype: str,
        length: str,
        package: str,
        request: str,
    ) -> Dict:
        return SapCommand(
            action="criar_dominio",
            vbs_url=VBS_URLS["criar_dominio"],
            args=[name, text, datatype, length, package, request],
        ).to_dict()

    def build_elemento_payload(
        self,
        name: str,
        text: str,
        domain: str,
        package: str,
        request: str,
    ) -> Dict:
        return SapCommand(
            action="criar_elemento",
            vbs_url=VBS_URLS["criar_elemento"],
            args=[name, text, domain, package, request],
        ).to_dict()

    def build_programa_payload(
        self,
        name: str,
        title: str,
        prog_type: str,
        prog_status: str,
        application: str,
        auth_group: str,
        package: str,
        request: str,
        flow: str,
        source_code: str = "",
    ) -> Dict:
        """
        Monta o payload para criacao de programa ABAP via SE38.

        Ordem dos args para alinhar com o VBS:
          args[0] = NomeProg (programName)
          args[1] = Pacote (packageName)
          args[2] = Request (requestId)
          args[3] = Titulo (titulo)
          args[4] = "" (não usado)
          args[5] = "" (não usado)
          args[6] = "" (não usado)
          args[7] = SourceCode (sourceCode)
        """
        encoded_code = source_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
        return SapCommand(
            action="criar_programa",
            vbs_url=VBS_URLS["criar_programa"],
            args=[
                name,        # programName
                package,     # packageName
                request,     # requestId
                title,       # titulo
                "",         # não usado
                "",         # não usado
                "",         # não usado
                encoded_code # sourceCode
            ],
        ).to_dict()

    def build_buscar_payload(self, name: str) -> Dict:
        """Busca codigo fonte de programa existente no SAP. args[0]=programName."""
        return SapCommand(
            action="buscar_programa",
            vbs_url=VBS_URLS["buscar_programa"],
            args=[name],
        ).to_dict()

    def build_read_file_v1(self, file_name: str) -> Dict:
        """
        Busca codigo fonte usando arquitetura V1 com query.read.file.
        Retorna payload + routing_key para publish no exchange topic.
        """
        event = SapEventV1(file_name=file_name)
        return {
            "payload": event.to_dict(),
            "routing_key": "usiminas.req.query.read.file.v1",
            "correlation_id": event.correlation_id,
        }

    def build_fix_v1_payload(self, name: str, code: str) -> Dict:
        """Edita código de programa ABAP existente via rota V1 (command.fix.v1)."""
        import uuid
        correlation_id = str(uuid.uuid4())
        encoded = code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
        return {
            "payload": {
                "fileName":     name,
                "category":     "PROGRAM",
                "functionGroup": "",
                "fileContent":  encoded,
                "contentType":  "MODIFIED",
                "correlationId": correlation_id,
                "replyTo":      QUEUE_RESPONSES,
            },
            "routing_key":    "usiminas.req.command.fix.v1",
            "correlation_id": correlation_id,
        }

    def build_buscar_arquivo_payload(self, filename: str) -> Dict:
        """Busca arquivo .txt via buscaArquivo.vbs. args[0]=fileName."""
        return SapCommand(
            action="buscar_arquivo",
            vbs_url=VBS_URLS["buscar_arquivo"],
            args=[filename],
        ).to_dict()

    def build_editar_payload(
        self,
        name: str,
        source_code: str,
        package: str = "",
        request: str = "",
    ) -> Dict:
        """
        Edita codigo de programa ABAP existente.
        args[0]=programName  args[1]=sourceCode  args[2]=package  args[3]=request
        """
        encoded = source_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
        return SapCommand(
            action="editar_programa",
            vbs_url=VBS_URLS["editar_programa"],
            args=[name, encoded, package, request],
        ).to_dict()

    def build_fields_spec(self, items: List[Dict]) -> str:
        """Serializa a lista de campos no formato esperado pelo VBS."""
        return ";".join(
            f"{it['field']}|{it['key']}|{it['notnull']}|{it['element_data']}"
            f"|{it.get('ref_table', '')}|{it.get('ref_field', '')}"
            for it in items
        )

    def build_message_chain(
        self,
        items: List[Dict],
        table_name: str,
        table_text: str,
        tabart: str,
        tabkat: str,
        package: str,
        request: str,
    ) -> List[Dict]:
        """
        Gera a cadeia completa de mensagens:
        dominios unicos -> elementos unicos -> tabela.
        Ignora MANDT e campos sem prefixo Z.
        """
        messages: List[Dict] = []
        done_dom: Set[str] = set()
        done_elem: Set[str] = set()

        for item in items:
            el = item["element_data"]
            if el == "MANDT" or not el.startswith("Z"):
                continue

            dom = item.get("domain_data", "").strip().upper() or el
            domtype = item.get("domtype", "").strip().upper() or DEFAULT_DOMAIN_DATATYPE
            domlen = item.get("domlen", "").strip() or DEFAULT_DOMAIN_LENGTH
            desc = item.get("desc", "").strip()
            field = item.get("field", "").strip().upper()
            dom_text = (
                desc or field.replace("_", " ").strip() or el.replace("_", " ").strip()
            )[:30]
            elem_text = desc or el

            if dom not in done_dom:
                messages.append(
                    self.build_dominio_payload(dom, dom_text, domtype, domlen, package, request)
                )
                done_dom.add(dom)

            if el not in done_elem:
                messages.append(
                    self.build_elemento_payload(el, elem_text, dom, package, request)
                )
                done_elem.add(el)

        fs = self.build_fields_spec(items)
        messages.append(
            SapCommand(
                action="criar_tabela",
                vbs_url=VBS_URLS["criar_tabela"],
                args=[table_name, table_text, package, request, tabart, tabkat, fs, "A"],
            ).to_dict()
        )

        return messages
