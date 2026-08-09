"""Owner-controlled authority ceiling for CLI requests.

The plan/command policy is a request. This module supplies the independent
maximum policy and live-backend grants that the CLI is allowed to exercise.
There is intentionally no command that writes this file.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import PolicyError, ValidationError
from .permissions import PROFILE_NAMES


AUTHORITY_SCHEMA = "white-collar.authority/v1"
AUTHORITY_APPS = ("word", "slides", "mail")
POLICY_RANK = {"read-only": 0, "review": 1, "edit": 2}
SUPPORTED_BACKENDS = frozenset({"word:com", "slides:com", "mail:com"})


@dataclass(frozen=True)
class Authority:
    """Maximum policy and backend grants approved outside an agent request."""

    maximum_policy: dict[str, str]
    enabled_backends: frozenset[str]
    source: str

    @classmethod
    def default(cls, *, source: str = "default") -> "Authority":
        return cls(
            maximum_policy={app: "read-only" for app in AUTHORITY_APPS},
            # Rendering and inspection through the existing Office backends are
            # available; Outlook remains explicitly disabled.
            enabled_backends=frozenset({"word:com", "slides:com"}),
            source=source,
        )

    @classmethod
    def for_testing(cls) -> "Authority":
        return cls(
            maximum_policy={"word": "edit", "slides": "edit", "mail": "review"},
            enabled_backends=SUPPORTED_BACKENDS,
            source="test",
        )

    @classmethod
    def from_dict(cls, raw: Any, *, source: str) -> "Authority":
        if not isinstance(raw, dict):
            raise ValidationError("authority must be a JSON object", details={"source": source})
        required = {"schema", "policies", "backends"}
        unknown = set(raw) - required
        missing = required - set(raw)
        if missing:
            raise ValidationError(
                "authority is missing required field(s)",
                details={"missing": sorted(missing), "source": source},
            )
        if unknown:
            raise ValidationError(
                "authority has unknown field(s)",
                details={"unknown": sorted(unknown), "source": source},
            )
        if raw["schema"] != AUTHORITY_SCHEMA:
            raise ValidationError(
                f"authority schema must be {AUTHORITY_SCHEMA!r}",
                details={"source": source},
            )
        policies = raw["policies"]
        if not isinstance(policies, dict):
            raise ValidationError("authority.policies must be an object")
        unknown_apps = set(policies) - set(AUTHORITY_APPS)
        if unknown_apps:
            raise ValidationError("authority.policies contains unknown app(s)", details={"apps": sorted(unknown_apps)})
        maximum = {app: "read-only" for app in AUTHORITY_APPS}
        for app, policy in policies.items():
            if policy not in PROFILE_NAMES:
                raise ValidationError(
                    "authority policy is not a supported profile",
                    details={"app": app, "policy": policy, "allowed": list(PROFILE_NAMES)},
                )
            maximum[app] = policy
        backends = raw["backends"]
        if not isinstance(backends, list) or not all(isinstance(item, str) for item in backends):
            raise ValidationError("authority.backends must be an array of strings")
        unsupported = set(backends) - SUPPORTED_BACKENDS
        if unsupported:
            raise ValidationError("authority contains unsupported backend(s)", details={"backends": sorted(unsupported)})
        return cls(maximum_policy=maximum, enabled_backends=frozenset(backends), source=source)

    def require_policy(self, app: str, requested: str) -> None:
        if app not in AUTHORITY_APPS:
            raise PolicyError("authority does not recognize app", details={"app": app})
        if requested not in POLICY_RANK:
            raise PolicyError("authority does not recognize policy", details={"policy": requested})
        maximum = self.maximum_policy.get(app, "read-only")
        if POLICY_RANK[requested] > POLICY_RANK[maximum]:
            raise PolicyError(
                "requested policy exceeds owner authority",
                details={
                    "app": app,
                    "requested_policy": requested,
                    "maximum_policy": maximum,
                    "authority_source": self.source,
                },
            )

    def require_backend(self, app: str, backend: str) -> None:
        if backend == "local":
            return
        grant = f"{app}:{backend}"
        if grant not in self.enabled_backends:
            raise PolicyError(
                "backend is disabled by owner authority",
                details={"backend": grant, "authority_source": self.source},
            )

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema": AUTHORITY_SCHEMA,
            "source": self.source,
            "maximum_policy": dict(self.maximum_policy),
            "enabled_backends": sorted(self.enabled_backends),
        }


def authority_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "white-collar" / "authority.json"
    return Path.home() / ".config" / "white-collar" / "authority.json"


def load_authority(path: Path | None = None) -> Authority:
    configured = path or authority_path()
    if not configured.exists():
        return Authority.default(source=f"default:{configured}")
    try:
        raw = json.loads(configured.read_text(encoding="utf-8"))
    except OSError as exc:
        raise ValidationError("cannot read authority file", details={"path": str(configured), "reason": str(exc)}) from exc
    except json.JSONDecodeError as exc:
        raise ValidationError(
            "authority file is not valid JSON",
            details={"path": str(configured), "line": exc.lineno, "column": exc.colno},
        ) from exc
    return Authority.from_dict(raw, source=str(configured))
