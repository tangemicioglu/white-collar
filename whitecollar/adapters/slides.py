from __future__ import annotations

from pathlib import Path
from typing import Any, Protocol

from ..errors import BackendUnavailableError
from ..models import Plan


class SlidesAdapter(Protocol):
    def inspect(self, target: Path) -> dict[str, Any]: ...

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]: ...


class UnavailableSlidesAdapter:
    def inspect(self, target: Path) -> dict[str, Any]:
        raise BackendUnavailableError("slides")

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
        raise BackendUnavailableError("slides")
