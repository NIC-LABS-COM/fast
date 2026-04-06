"""
API bridge entre web UI (pywebview) e servicos Python.
Expoe metodos que JavaScript chama via window.pywebview.api.
"""
import json
import logging
import threading
import uuid
from typing import Any, Dict, List

from core.config import (
    DEFAULT_PACKAGE, DEFAULT_REQUEST,
    TABART_OPTIONS, TABKAT_OPTIONS, QUEUE_RESPONSES,
    PROG_TYPE_OPTIONS, PROG_FLOW_OPTIONS,
    TYPES_REQUIRING_LENGTH,
)
from services.ai_service import AIService
from services.publisher_service import PublisherService
from services.rabbitmq_service import RabbitMQService

logger = logging.getLogger("sap_publisher")


class Api:
    """Python API exposta ao JavaScript via pywebview."""

    def __init__(self) -> None:
        self._publisher = PublisherService()
        self._rabbitmq = RabbitMQService()
        self._ai = AIService()
        self._window = None
        self._pending_responses: List[Any] = []
        self._lock = threading.Lock()

    def _set_window(self, window) -> None:
        self._window = window
        self._rabbitmq.start_listener(self._on_response)

    def _on_response(self, response: Any) -> None:
        """Called from pika thread — just queues the response."""
        with self._lock:
            self._pending_responses.append(response)

    def get_pending_responses(self) -> List[Any]:
        """Called from JS via polling — returns and clears queued responses."""
        with self._lock:
            items = self._pending_responses[:]
            self._pending_responses.clear()
        return items

    # ── Config ──────────────────────────────────────────────────────
    def get_config(self) -> Dict:
        return {
            "defaultPackage": DEFAULT_PACKAGE,
            "defaultRequest": DEFAULT_REQUEST,
            "tabartOptions": TABART_OPTIONS,
            "tabkatOptions": TABKAT_OPTIONS,
            "progTypeOptions": PROG_TYPE_OPTIONS,
            "progFlowOptions": PROG_FLOW_OPTIONS,
            "typesRequiringLength": list(TYPES_REQUIRING_LENGTH),
        }

    # ── Domain ──────────────────────────────────────────────────────
    def publish_domain(self, name, text, datatype, length, package, request):
        try:
            payload = self._publisher.build_dominio_payload(
                name.strip().upper(), text.strip(), datatype.strip().upper(),
                length.strip(), package.strip(), request.strip().upper(),
            )
            ok = self._rabbitmq.publish(payload)
            return {"success": ok}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── Element ─────────────────────────────────────────────────────
    def publish_element(self, name, text, domain, package, request):
        try:
            payload = self._publisher.build_elemento_payload(
                name.strip().upper(), text.strip(), domain.strip().upper(),
                package.strip(), request.strip().upper(),
            )
            ok = self._rabbitmq.publish(payload)
            return {"success": ok}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── Table ───────────────────────────────────────────────────────
    def publish_table(self, table_name, table_text, package, request, tabart, tabkat, fields):
        try:
            messages = self._publisher.build_message_chain(
                fields, table_name.strip().upper(), table_text.strip(),
                tabart.strip(), tabkat.strip(),
                package.strip(), request.strip().upper(),
            )
            if len(messages) == 1:
                ok = self._rabbitmq.publish(messages[0])
                return {"success": ok, "count": 1}
            ok_count, fail_count = self._rabbitmq.publish_batch(messages)
            return {"success": fail_count == 0, "count": ok_count, "failed": fail_count}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def generate_table_ai(self, prompt, state):
        try:
            result = self._ai.call(prompt, state)
            return {"success": True, "data": result}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── SE38 ────────────────────────────────────────────────────────
    def publish_program(self, name, title, prog_type, package, request, flow, source_code):
        try:
            payload = self._publisher.build_programa_payload(
                name.strip().upper(), title.strip(), prog_type, "", "", "",
                package.strip(), request.strip().upper(), flow, source_code,
            )
            ok = self._rabbitmq.publish(payload)
            return {"success": ok}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def publish_fix(self, name, code):
        try:
            v1_data = self._publisher.build_fix_v1_payload(name.strip().upper(), code)
            ok = self._rabbitmq.publish_v1(v1_data["routing_key"], v1_data["payload"])
            return {"success": ok}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_program(self, name):
        try:
            v1_data = self._publisher.build_read_file_v1(name.strip().upper())
            ok = self._rabbitmq.publish_v1(v1_data["routing_key"], v1_data["payload"])
            return {"success": ok}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def generate_program_ai(self, prompt, state):
        try:
            result = self._ai.call_se38(prompt, state)
            return {"success": True, "data": result}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── Queries ─────────────────────────────────────────────────────
    def _publish_query(self, routing_key, payload_extra=None):
        try:
            cid = str(uuid.uuid4())
            payload = {"correlationId": cid, "replyTo": QUEUE_RESPONSES}
            if payload_extra:
                payload.update(payload_extra)
            ok = self._rabbitmq.publish_v1(routing_key, payload)
            return {"success": ok, "correlationId": cid}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def query_requests(self):
        return self._publish_query("usiminas.req.query.requests.v1", {"QueryFilterString": ""})

    def query_reports(self):
        return self._publish_query("usiminas.req.query.all.files.v1")

    def query_packages(self):
        return self._publish_query("usiminas.req.query.all.packages.v1")

    def query_versions_metadata(self, file_name, category):
        return self._publish_query("usiminas.req.query.versions.metadata.v1", {
            "fileName": file_name, "category": category,
        })

    def query_file_category(self, file_name):
        return self._publish_query("usiminas.req.query.file.category.v1", {"fileName": file_name})

    def query_request_files(self, requests_csv):
        requests_list = [r.strip() for r in requests_csv.split(",") if r.strip()]
        return self._publish_query("usiminas.req.query.request.files.v1", {"requests": requests_list})

    def query_request_description(self, request_id):
        return self._publish_query("usiminas.req.query.request.description.v1", {"requestId": request_id})

    def query_read_from_version(self, file_name, category, version_id):
        return self._publish_query("usiminas.req.query.read.from.version.v1", {
            "fileName": file_name, "category": category, "versionId": version_id,
        })

    def query_search(self, file_filter, package_filter):
        return self._publish_query("usiminas.req.query.search.v1", {
            "fileFilter": file_filter, "packageFilter": package_filter,
        })
