from __future__ import annotations

from pathlib import Path
from typing import Any, Protocol


class WordAdapter(Protocol):
    def inspect(self, target: Path, *, render_dir: Path | None = None) -> dict[str, Any]: ...

    def apply(self, plan: Any, *, dry_run: bool) -> dict[str, Any]: ...


class SlidesAdapter(Protocol):
    def inspect(self, target: Path, *, render_dir: Path | None = None) -> dict[str, Any]: ...

    def apply(self, plan: Any, *, dry_run: bool) -> dict[str, Any]: ...


class MailAdapter(Protocol):
    def search(self, query: str, *, limit: int, folder: str = "Inbox") -> list[dict[str, Any]]: ...

    def read(self, message_id: str, *, include_body: bool = False) -> dict[str, Any]: ...

    def apply(self, plan: Any, *, dry_run: bool) -> dict[str, Any]: ...
