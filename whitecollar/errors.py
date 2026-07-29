from __future__ import annotations

from typing import Any


class WhiteCollarError(Exception):
    """An expected error that can be returned as structured JSON."""

    def __init__(self, code: str, message: str, *, details: dict[str, Any] | None = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or {}


class ValidationError(WhiteCollarError):
    def __init__(self, message: str, *, details: dict[str, Any] | None = None):
        super().__init__("validation_error", message, details=details)


class PolicyError(WhiteCollarError):
    def __init__(self, message: str, *, details: dict[str, Any] | None = None):
        super().__init__("policy_denied", message, details=details)


class BackendUnavailableError(WhiteCollarError):
    def __init__(self, app: str):
        super().__init__(
            "backend_unavailable",
            f"no {app} backend is configured",
            details={"app": app, "hint": "inject or configure an app-specific backend"},
        )
