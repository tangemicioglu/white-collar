from __future__ import annotations

from typing import Any, Protocol

from ..errors import BackendUnavailableError


class MailAdapter(Protocol):
    def search(self, query: str, *, limit: int) -> list[dict[str, Any]]: ...

    def read(self, message_id: str) -> dict[str, Any]: ...


class UnavailableMailAdapter:
    def search(self, query: str, *, limit: int) -> list[dict[str, Any]]:
        raise BackendUnavailableError("mail")

    def read(self, message_id: str) -> dict[str, Any]:
        raise BackendUnavailableError("mail")
