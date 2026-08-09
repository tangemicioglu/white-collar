"""User-mediated capability grants backed by protected OS credential storage.

The agent may request a capability in a plan or command. It may not create the
grant that authorizes that request. Owner grants are stored in Windows
Credential Manager on Windows and in an installed OS-keyring backend
elsewhere; there is deliberately no JSON-file fallback.
"""

from __future__ import annotations

import ctypes
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

from .errors import PolicyError, ValidationError
from .permissions import CAPABILITIES, PROFILE_CAPABILITIES, PROFILE_NAMES


AUTHORITY_SCHEMA = "white-collar.authority/v2"
GRANT_SCHEMA = "white-collar.grant/v1"
AUTHORITY_APPS = ("word", "slides", "mail")
BACKEND_NAMES = ("local", "com")
POLICY_RANK = {"read-only": 0, "review": 1, "edit": 2}
CREDENTIAL_TARGET = "white-collar.owner-grant.v1"
HUMAN_PERMISSION_NOTICE = (
    "Permission changes are human-owner-only. An agent must stop and ask a human "
    "to approve this grant in their own interactive terminal; it must not create, "
    "edit, confirm, or retry the grant."
)
HUMAN_CONFIRMATION_PHRASE = "I AM THE HUMAN OWNER"


class CredentialStore(Protocol):
    """Small protected-store boundary; implementations never expose a file path."""

    def read(self, target: str = CREDENTIAL_TARGET) -> bytes | None: ...

    def write(self, value: bytes, target: str = CREDENTIAL_TARGET) -> None: ...

    def delete(self, target: str = CREDENTIAL_TARGET) -> None: ...


class MemoryCredentialStore:
    """Test-only store. Production code uses an OS-protected implementation."""

    def __init__(self, value: bytes | None = None):
        self.value = value

    def read(self, target: str = CREDENTIAL_TARGET) -> bytes | None:
        return self.value

    def write(self, value: bytes, target: str = CREDENTIAL_TARGET) -> None:
        self.value = bytes(value)

    def delete(self, target: str = CREDENTIAL_TARGET) -> None:
        self.value = None


if os.name == "nt":

    class _FILETIME(ctypes.Structure):
        _fields_ = [("dwLowDateTime", ctypes.c_uint32), ("dwHighDateTime", ctypes.c_uint32)]

    class _CREDENTIAL_ATTRIBUTEW(ctypes.Structure):
        _fields_ = [
            ("Keyword", ctypes.c_wchar_p),
            ("Flags", ctypes.c_uint32),
            ("ValueSize", ctypes.c_uint32),
            ("Value", ctypes.POINTER(ctypes.c_ubyte)),
        ]

    class _CREDENTIALW(ctypes.Structure):
        _fields_ = [
            ("Flags", ctypes.c_uint32),
            ("Type", ctypes.c_uint32),
            ("TargetName", ctypes.c_wchar_p),
            ("Comment", ctypes.c_wchar_p),
            ("LastWritten", _FILETIME),
            ("CredentialBlobSize", ctypes.c_uint32),
            ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
            ("Persist", ctypes.c_uint32),
            ("AttributeCount", ctypes.c_uint32),
            ("Attributes", ctypes.POINTER(_CREDENTIAL_ATTRIBUTEW)),
            ("TargetAlias", ctypes.c_wchar_p),
            ("UserName", ctypes.c_wchar_p),
        ]


class WindowsCredentialStore:
    """Native Windows Credential Manager backend.

    Credential Manager protects the blob with the logged-in Windows user's
    credential boundary. The grant is not read from an agent-authored file.
    """

    def __init__(self) -> None:
        if os.name != "nt":
            raise RuntimeError("WindowsCredentialStore is available only on Windows")
        self._dll = ctypes.WinDLL("Advapi32.dll", use_last_error=True)
        self._read = self._dll.CredReadW
        self._read.argtypes = [
            ctypes.c_wchar_p,
            ctypes.c_uint32,
            ctypes.c_uint32,
            ctypes.POINTER(ctypes.POINTER(_CREDENTIALW)),
        ]
        self._read.restype = ctypes.c_bool
        self._write = self._dll.CredWriteW
        self._write.argtypes = [ctypes.POINTER(_CREDENTIALW), ctypes.c_uint32]
        self._write.restype = ctypes.c_bool
        self._delete = self._dll.CredDeleteW
        self._delete.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32]
        self._delete.restype = ctypes.c_bool
        self._free = self._dll.CredFree
        self._free.argtypes = [ctypes.c_void_p]
        self._free.restype = None

    @staticmethod
    def _error(operation: str) -> OSError:
        code = ctypes.get_last_error()
        return OSError(code, f"{operation} failed with Windows error {code}")

    def read(self, target: str = CREDENTIAL_TARGET) -> bytes | None:
        credential = ctypes.POINTER(_CREDENTIALW)()
        if not self._read(target, 1, 0, ctypes.byref(credential)):
            if ctypes.get_last_error() == 1168:  # ERROR_NOT_FOUND
                return None
            raise self._error("CredReadW")
        try:
            value = credential.contents
            if not value.CredentialBlobSize:
                return b""
            return ctypes.string_at(value.CredentialBlob, value.CredentialBlobSize)
        finally:
            self._free(credential)

    def write(self, value: bytes, target: str = CREDENTIAL_TARGET) -> None:
        blob = (ctypes.c_ubyte * len(value)).from_buffer_copy(value)
        credential = _CREDENTIALW()
        credential.Type = 1  # CRED_TYPE_GENERIC
        credential.TargetName = target
        credential.CredentialBlobSize = len(value)
        credential.CredentialBlob = ctypes.cast(blob, ctypes.POINTER(ctypes.c_ubyte))
        credential.Persist = 2  # CRED_PERSIST_LOCAL_MACHINE
        if not self._write(ctypes.byref(credential), 0):
            raise self._error("CredWriteW")

    def delete(self, target: str = CREDENTIAL_TARGET) -> None:
        if not self._delete(target, 1, 0):
            if ctypes.get_last_error() == 1168:  # ERROR_NOT_FOUND
                return
            raise self._error("CredDeleteW")


class KeyringCredentialStore:
    """Optional cross-platform OS-keyring backend; never falls back to a file."""

    def __init__(self) -> None:
        try:
            import keyring  # type: ignore
        except ImportError as exc:  # pragma: no cover - depends on environment
            raise RuntimeError("install the keyring package for protected grant storage") from exc
        self._keyring = keyring

    def read(self, target: str = CREDENTIAL_TARGET) -> bytes | None:
        value = self._keyring.get_password("white-collar", target)
        return value.encode("utf-8") if value is not None else None

    def write(self, value: bytes, target: str = CREDENTIAL_TARGET) -> None:
        self._keyring.set_password("white-collar", target, value.decode("utf-8"))

    def delete(self, target: str = CREDENTIAL_TARGET) -> None:
        try:
            self._keyring.delete_password("white-collar", target)
        except Exception as exc:  # keyring backends use different not-found errors
            if "not found" not in str(exc).lower() and "no password" not in str(exc).lower():
                raise


def default_credential_store() -> CredentialStore | None:
    if os.name == "nt":
        return WindowsCredentialStore()
    try:
        return KeyringCredentialStore()
    except RuntimeError:
        return None


@dataclass(frozen=True)
class Grant:
    app: str
    backend: str
    policy: str
    capabilities: tuple[str, ...]
    targets: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "app": self.app,
            "backend": self.backend,
            "policy": self.policy,
            "capabilities": list(self.capabilities),
            "targets": list(self.targets),
        }


def _builtin_grants() -> tuple[Grant, ...]:
    return (
        Grant("word", "local", "read-only", ("word.read", "word.render"), ("*",)),
        Grant("word", "com", "read-only", ("word.read", "word.render"), ("*",)),
        Grant("slides", "local", "read-only", ("slides.read", "slides.render"), ("*",)),
        Grant("slides", "com", "read-only", ("slides.read", "slides.render"), ("*",)),
        Grant("mail", "local", "read-only", ("mail.metadata.read",), ("mailbox",)),
    )


@dataclass(frozen=True)
class Authority:
    """Effective authority: safe built-ins plus owner grants from the OS store."""

    owner_grants: tuple[Grant, ...] = ()
    source: str = "default"

    @property
    def grants(self) -> tuple[Grant, ...]:
        return _builtin_grants() + self.owner_grants

    @classmethod
    def default(cls, *, source: str = "default") -> "Authority":
        return cls((), source)

    @classmethod
    def for_testing(cls) -> "Authority":
        from .permissions import CAPABILITIES

        all_caps = tuple(CAPABILITIES)
        return cls(
            tuple(
                Grant(app, backend, "edit", all_caps, ("*",))
                for app in AUTHORITY_APPS
                for backend in BACKEND_NAMES
            ),
            "test",
        )

    @classmethod
    def from_grant_dict(cls, raw: Any, *, source: str) -> "Authority":
        if not isinstance(raw, dict):
            raise ValidationError("protected grant payload must be a JSON object", details={"source": source})
        required = {"schema", "grants"}
        unknown = set(raw) - required
        missing = required - set(raw)
        if missing:
            raise ValidationError("protected grant payload is missing required field(s)", details={"missing": sorted(missing)})
        if unknown:
            raise ValidationError("protected grant payload has unknown field(s)", details={"unknown": sorted(unknown)})
        if raw["schema"] != GRANT_SCHEMA:
            raise ValidationError(f"protected grant schema must be {GRANT_SCHEMA!r}")
        grants = raw["grants"]
        if not isinstance(grants, list):
            raise ValidationError("protected grant payload grants must be an array")
        parsed = tuple(_parse_grant(item, index) for index, item in enumerate(grants))
        return cls(parsed, source)

    def to_dict(self) -> dict[str, Any]:
        """Return safe status data; no credential blob or storage path is exposed."""

        return {
            "schema": AUTHORITY_SCHEMA,
            "source": self.source,
            "built_in": [grant.to_dict() for grant in _builtin_grants()],
            "owner_grants": [grant.to_dict() for grant in self.owner_grants],
            "human_action_required_for_changes": True,
        }

    def grant_payload(self) -> dict[str, Any]:
        return {"schema": GRANT_SCHEMA, "grants": [grant.to_dict() for grant in self.owner_grants]}

    def require_backend(self, app: str, backend: str) -> None:
        if backend not in BACKEND_NAMES:
            raise PolicyError("unknown backend", details={"app": app, "backend": backend})
        if not any(grant.app == app and grant.backend == backend for grant in self.grants):
            raise _permission_denied(app, backend, "read-only", (), None)

    def has_capability(self, app: str, capability: str, *, policy: str | None = None) -> bool:
        for backend in BACKEND_NAMES:
            for grant in self.grants:
                if grant.app != app or capability not in grant.capabilities:
                    continue
                if policy is None or POLICY_RANK[grant.policy] >= POLICY_RANK[policy]:
                    return True
        return False

    def require_policy(
        self,
        app: str,
        requested: str,
        *,
        backend: str = "local",
        target: str | None = None,
        capabilities: tuple[str, ...] = (),
    ) -> None:
        if requested not in POLICY_RANK:
            raise PolicyError("authority does not recognize policy", details={"policy": requested})
        if not capabilities:
            capabilities = tuple(
                name for name, spec in CAPABILITIES.items()
                if spec.app == app and name in PROFILE_CAPABILITIES[requested]
            )
        self.require_access(app, backend, requested, target, capabilities)

    def require_access(
        self,
        app: str,
        backend: str,
        requested: str,
        target: str | None,
        capabilities: tuple[str, ...] | list[str] | set[str],
    ) -> None:
        if requested not in POLICY_RANK:
            raise PolicyError("authority does not recognize policy", details={"policy": requested})
        needed = tuple(sorted(set(capabilities)))
        for grant in self.grants:
            if grant.app != app or grant.backend != backend:
                continue
            if POLICY_RANK[grant.policy] < POLICY_RANK[requested]:
                continue
            if not all(capability in grant.capabilities for capability in needed):
                continue
            if target is not None and not any(_target_matches(app, target, allowed) for allowed in grant.targets):
                continue
            return
        raise _permission_denied(app, backend, requested, needed, target)


def _permission_denied(
    app: str,
    backend: str,
    policy: str,
    capabilities: tuple[str, ...] | list[str] | set[str],
    target: str | None,
) -> PolicyError:
    command = f"white-collar permissions grant --app {app} --backend {backend} --policy {policy}"
    if target:
        command += f' --target "{target}"'
    return PolicyError(
        "requested access has not been approved by the human owner",
        details={
            "app": app,
            "backend": f"{app}:{backend}",
            "backend_kind": backend,
            "requested_policy": policy,
            "capabilities": sorted(capabilities),
            "target": target,
            "human_action_required": True,
            "agent_instruction": HUMAN_PERMISSION_NOTICE,
            "human_only_command": command,
        },
    )


def _target_matches(app: str, requested: str, allowed: str) -> bool:
    if allowed == "*":
        return True
    if app in {"word", "slides"}:
        try:
            return Path(requested).resolve() == Path(allowed).resolve()
        except (OSError, RuntimeError):
            return requested == allowed
    return requested == allowed


def _parse_grant(raw: Any, index: int) -> Grant:
    context = f"grants[{index}]"
    if not isinstance(raw, dict):
        raise ValidationError(f"{context} must be an object")
    required = {"app", "backend", "policy", "capabilities", "targets"}
    unknown = set(raw) - required
    missing = required - set(raw)
    if missing or unknown:
        raise ValidationError(f"{context} has invalid fields", details={"missing": sorted(missing), "unknown": sorted(unknown)})
    app = raw["app"]
    backend = raw["backend"]
    policy = raw["policy"]
    capabilities = raw["capabilities"]
    targets = raw["targets"]
    if app not in AUTHORITY_APPS or backend not in BACKEND_NAMES or policy not in PROFILE_NAMES:
        raise ValidationError(f"{context} has an unsupported app, backend, or policy")
    if not isinstance(capabilities, list) or not capabilities or not all(isinstance(item, str) for item in capabilities):
        raise ValidationError(f"{context}.capabilities must be a non-empty string array")
    if not isinstance(targets, list) or not targets or not all(isinstance(item, str) and item for item in targets):
        raise ValidationError(f"{context}.targets must be a non-empty string array")
    from .permissions import CAPABILITIES

    if any(item not in CAPABILITIES or CAPABILITIES[item].app != app for item in capabilities):
        raise ValidationError(f"{context}.capabilities contains an unknown or cross-app capability")
    if any(item not in PROFILE_CAPABILITIES[policy] for item in capabilities):
        raise ValidationError(f"{context}.capabilities exceeds its policy profile")
    return Grant(app, backend, policy, tuple(sorted(set(capabilities))), tuple(targets))


def _all_profile_capabilities(app: str, policy: str) -> tuple[str, ...]:
    from .permissions import CAPABILITIES

    return tuple(sorted(name for name, spec in CAPABILITIES.items() if spec.app == app and name in PROFILE_CAPABILITIES[policy]))


def make_grant(
    *,
    app: str,
    backend: str,
    policy: str,
    targets: list[str],
    capabilities: list[str] | None = None,
) -> Grant:
    if app not in AUTHORITY_APPS:
        raise ValidationError("unknown grant app", details={"app": app})
    if backend not in BACKEND_NAMES:
        raise ValidationError("unknown grant backend", details={"backend": backend})
    if policy not in PROFILE_NAMES:
        raise ValidationError("unknown grant policy", details={"policy": policy})
    if not targets or any(not target for target in targets):
        raise ValidationError("grant requires at least one non-empty target")
    selected = tuple(sorted(set(capabilities or _all_profile_capabilities(app, policy))))
    grant = Grant(app, backend, policy, selected, tuple(_normalize_target(app, target) for target in targets))
    _parse_grant(grant.to_dict(), 0)
    return grant


def _normalize_target(app: str, target: str) -> str:
    if app in {"word", "slides"} and target != "*":
        path = Path(target)
        if not path.is_absolute():
            raise ValidationError("file grant targets must be absolute paths", details={"target": target})
        return str(path.resolve())
    return target


def save_grant(authority: Authority, grant: Grant, store: CredentialStore | None = None) -> Authority:
    active_store = store or default_credential_store()
    if active_store is None:
        raise ValidationError(
            "protected OS credential storage is unavailable",
            details={"human_action_required": True, "hint": "install an OS keyring backend; no plaintext file fallback is supported"},
        )
    existing = list(authority.owner_grants)
    if grant not in existing:
        existing.append(grant)
    updated = Authority(tuple(existing), source="os-credential-store")
    _write_authority(updated, active_store)
    return updated


def revoke_grant(authority: Authority, grant: Grant, store: CredentialStore | None = None) -> Authority:
    active_store = store or default_credential_store()
    if active_store is None:
        raise ValidationError("protected OS credential storage is unavailable", details={"human_action_required": True})
    updated_items: list[Grant] = []
    requested_capabilities = set(grant.capabilities)
    requested_targets = set(grant.targets)
    for item in authority.owner_grants:
        same_scope = (
            item.app == grant.app
            and item.backend == grant.backend
            and item.policy == grant.policy
            and set(item.targets) == requested_targets
            and requested_capabilities.issubset(item.capabilities)
        )
        if not same_scope:
            updated_items.append(item)
            continue
        remaining = tuple(sorted(set(item.capabilities) - requested_capabilities))
        if remaining:
            updated_items.append(Grant(item.app, item.backend, item.policy, remaining, item.targets))
    updated = Authority(tuple(updated_items), source="os-credential-store")
    _write_authority(updated, active_store)
    return updated


def revoke_all_grants(authority: Authority, store: CredentialStore | None = None) -> Authority:
    active_store = store or default_credential_store()
    if active_store is None:
        raise ValidationError("protected OS credential storage is unavailable", details={"human_action_required": True})
    updated = Authority((), source="os-credential-store")
    _write_authority(updated, active_store)
    return updated


def _write_authority(authority: Authority, store: CredentialStore) -> None:
    value = json.dumps(authority.grant_payload(), sort_keys=True, separators=(",", ":")).encode("utf-8")
    try:
        store.write(value)
    except Exception as exc:
        raise ValidationError(
            "cannot write protected owner grant",
            details={"human_action_required": True, "reason": str(exc)},
        ) from exc


def load_authority(
    path: Path | None = None,
    *,
    store: CredentialStore | None = None,
) -> Authority:
    """Load only from protected storage; ``path`` is rejected intentionally."""

    if path is not None:
        raise ValidationError(
            "authority files are not supported; owner grants must use protected OS credential storage",
            details={"human_action_required": True, "agent_instruction": HUMAN_PERMISSION_NOTICE},
        )
    active_store = store or default_credential_store()
    if active_store is None:
        return Authority.default(source="default:no-protected-store")
    try:
        value = active_store.read()
    except Exception as exc:
        raise ValidationError("cannot read protected owner grant", details={"reason": str(exc)}) from exc
    if value is None:
        return Authority.default(source="default:no-owner-grant")
    try:
        raw = json.loads(value.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValidationError("protected owner grant is corrupt", details={"reason": str(exc)}) from exc
    return Authority.from_grant_dict(raw, source="os-credential-store")
