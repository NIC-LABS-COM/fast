"""
Modelos de dados para comandos e respostas SAP via RabbitMQ.
"""
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import List, Optional

from core.config import QUEUE_RESPONSES


def _new_correlation_id() -> str:
    return str(uuid.uuid4())


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


@dataclass
class SapCommand:
    """Representa um comando enviado para o consumer SAP (formato legado)."""

    action: str
    vbs_url: str
    args: List[str] = field(default_factory=list)
    correlation_id: str = field(default_factory=_new_correlation_id)
    timestamp: str = field(default_factory=_now_iso)
    reply_to: str = QUEUE_RESPONSES

    def to_dict(self) -> dict:
        return {
            "correlationId": self.correlation_id,
            "action": self.action,
            "vbs_url": self.vbs_url,
            "args": self.args,
            "timestamp": self.timestamp,
            "replyTo": self.reply_to,
        }


@dataclass
class SapEventV1:
    """Representa um evento V1 — nova arquitetura orientada a eventos."""

    file_name: str
    content: str = ""
    correlation_id: str = field(default_factory=_new_correlation_id)
    reply_to: str = QUEUE_RESPONSES

    def to_dict(self) -> dict:
        return {
            "fileName": self.file_name,
            "content": self.content,
            "correlationId": self.correlation_id,
            "replyTo": self.reply_to,
        }


@dataclass
class SapResponse:
    """Representa uma resposta recebida do consumer SAP."""

    correlation_id: str
    action: str
    object_name: str
    status: str
    message: str
    timestamp: str = ""

    @classmethod
    def from_dict(cls, data: dict) -> "SapResponse":
        return cls(
            correlation_id=data.get("correlationId", ""),
            action=data.get("action", "?"),
            object_name=data.get("object_name", "?"),
            status=data.get("status", "?"),
            message=data.get("message", ""),
            timestamp=data.get("timestamp", ""),
        )
