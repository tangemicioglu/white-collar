from __future__ import annotations

from typing import Any, Protocol

from ..errors import BackendUnavailableError


class MailAdapter(Protocol):
    def search(self, query: str, *, limit: int, folder: str = "Inbox") -> list[dict[str, Any]]: ...

    def read(self, message_id: str, *, include_body: bool = False) -> dict[str, Any]: ...

    def apply(self, plan: Any, *, dry_run: bool) -> dict[str, Any]: ...


class UnavailableMailAdapter:
    def search(self, query: str, *, limit: int, folder: str = "Inbox") -> list[dict[str, Any]]:
        raise BackendUnavailableError("mail")

    def read(self, message_id: str, *, include_body: bool = False) -> dict[str, Any]:
        raise BackendUnavailableError("mail")

    def apply(self, plan: Any, *, dry_run: bool) -> dict[str, Any]:
        raise BackendUnavailableError("mail")
