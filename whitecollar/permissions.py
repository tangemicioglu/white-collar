"""Shared capability and permission decisions for the public CLI.

The permission vocabulary is deliberately smaller than the adapter vocabularies.
Plans can contain only finite semantic operations; this module maps those
operations to a handful of capabilities that are shared across applications.
It is an authorization layer for live Office actions, not an OAuth scope or a
generic COM security surface.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from .errors import PolicyError
from .mail_ops import MAIL_COM_OPERATIONS
from .slides_ops import SLIDES_COM_OPERATIONS, SLIDES_COM_READ_OPERATIONS
from .word_ops import WORD_COM_OPERATIONS, WORD_COM_READ_OPERATIONS


PERMISSIONS_SCHEMA = "white-collar.permissions/v1"
PROFILE_NAMES = ("read-only", "review", "edit", "send")
SETUP_POLICY_NAMES = ("disabled", "read-only", "review", "edit", "send")
SETUP_APP_POLICIES = {
    "word": ("disabled", "read-only", "review", "edit"),
    "slides": ("disabled", "read-only", "review", "edit"),
    "mail": SETUP_POLICY_NAMES,
}
SETUP_PRESETS = {
    "safe": {"word": "review", "slides": "review", "mail": "disabled"},
    "office-authoring": {"word": "edit", "slides": "edit", "mail": "disabled"},
    "outlook-review": {"mail": "review"},
    "outlook-send": {"mail": "send"},
}


@dataclass(frozen=True)
class Capability:
    name: str
    app: str
    description: str
    risk: str
    target: str


CAPABILITIES: dict[str, Capability] = {
    "word.read": Capability("word.read", "word", "Read Word document structure and text", "read", "file"),
    "word.render": Capability("word.render", "word", "Render Word pages to image files", "read", "file"),
    "word.capture": Capability("word.capture", "word", "Capture the visible Word window", "sensitive", "file"),
    "word.write.content": Capability("word.write.content", "word", "Change Word document content", "write", "file"),
    "word.write.comments": Capability("word.write.comments", "word", "Change Word comments", "write", "file"),
    "word.write.metadata": Capability("word.write.metadata", "word", "Change Word core properties", "write", "file"),
    "word.write.create": Capability("word.write.create", "word", "Create a new Word document", "write", "file"),
    "word.write.save_as": Capability("word.write.save_as", "word", "Create a new Word output file", "write", "file"),
    "word.write.in_place": Capability("word.write.in_place", "word", "Replace the Word target after snapshot", "destructive", "file"),
    "slides.read": Capability("slides.read", "slides", "Read PowerPoint presentation structure and text", "read", "file"),
    "slides.render": Capability("slides.render", "slides", "Render PowerPoint slides to image files", "read", "file"),
    "slides.capture": Capability("slides.capture", "slides", "Capture the visible PowerPoint window", "sensitive", "file"),
    "slides.write.content": Capability("slides.write.content", "slides", "Change PowerPoint slide content", "write", "file"),
    "slides.write.create": Capability("slides.write.create", "slides", "Create a new PowerPoint presentation", "write", "file"),
    "slides.write.save_as": Capability("slides.write.save_as", "slides", "Create a new PowerPoint output file", "write", "file"),
    "slides.write.in_place": Capability("slides.write.in_place", "slides", "Replace the PowerPoint target after snapshot", "destructive", "file"),
    "mail.metadata.read": Capability("mail.metadata.read", "mail", "Read mail headers and search metadata", "sensitive", "mailbox"),
    "mail.body.read": Capability("mail.body.read", "mail", "Read the body of a specific message", "sensitive", "message"),
    "mail.attachments.read": Capability("mail.attachments.read", "mail", "Read message attachments", "highly-sensitive", "message"),
    "mail.write.state": Capability("mail.write.state", "mail", "Mark an Outlook message read or unread", "write", "message"),
    "mail.write.organize": Capability("mail.write.organize", "mail", "Move or delete mail", "destructive", "message"),
    "mail.write.compose": Capability("mail.write.compose", "mail", "Create an Outlook draft", "write", "mailbox"),
    "mail.write.send": Capability("mail.write.send", "mail", "Send an existing Outlook draft", "destructive", "message"),
}


def _operation_capability(app: str, operation: str) -> str:
    if operation == "replace_text":
        return f"{app}.write.content"
    if app == "word":
        if operation == "word_live_create_document":
            return "word.write.create"
        if operation in WORD_COM_READ_OPERATIONS:
            return "word.read"
        if operation == "word_screen_capture":
            return "word.capture"
        if operation == "word_live_set_core_properties":
            return "word.write.metadata"
        if operation in {"word_live_add_comment", "word_live_reply_to_comment", "word_live_resolve_comment", "word_live_delete_comment"}:
            return "word.write.comments"
        if operation in WORD_COM_OPERATIONS:
            return "word.write.content"
    if app == "slides":
        if operation == "slides_live_create_presentation":
            return "slides.write.create"
        if operation in SLIDES_COM_READ_OPERATIONS:
            return "slides.read"
        if operation == "slides_screen_capture":
            return "slides.capture"
        if operation in SLIDES_COM_OPERATIONS:
            return "slides.write.content"
    if app == "mail" and operation in MAIL_COM_OPERATIONS:
        if operation in {"mail_live_mark_read", "mail_live_mark_unread"}:
            return "mail.write.state"
        if operation == "mail_live_create_draft":
            return "mail.write.compose"
        if operation == "mail_live_send":
            return "mail.write.send"
        return "mail.write.organize"
    raise PolicyError(
        "operation has no registered capability",
        details={"app": app, "operation": operation},
    )


OPERATION_CAPABILITIES = {
    (app, operation): _operation_capability(app, operation)
    for app, operations in (("word", WORD_COM_OPERATIONS), ("slides", SLIDES_COM_OPERATIONS))
    for operation in operations
}


_READ_ONLY = frozenset({
    "word.read",
    "word.render",
    "slides.read",
    "slides.render",
    "mail.metadata.read",
})
_REVIEW = _READ_ONLY | frozenset({
    "word.capture",
    "word.write.content",
    "word.write.comments",
    "word.write.metadata",
    "word.write.create",
    "word.write.save_as",
    "slides.capture",
    "slides.write.content",
    "slides.write.create",
    "slides.write.save_as",
    "mail.body.read",
    "mail.write.state",
})
_EDIT = _REVIEW | frozenset({
    "word.write.in_place",
    "slides.write.in_place",
    "mail.write.organize",
    "mail.write.compose",
})
_SEND = _EDIT | frozenset({"mail.write.send"})

PROFILE_CAPABILITIES = {
    "read-only": _READ_ONLY,
    "review": _REVIEW,
    "edit": _EDIT,
    "send": _SEND,
}


def setup_capabilities(app: str, policy: str) -> tuple[str, ...]:
    """Return the bounded capability bundle used by the human setup wizard."""

    if policy == "disabled":
        return ()
    if app not in SETUP_APP_POLICIES or policy not in SETUP_APP_POLICIES[app]:
        raise PolicyError(
            "policy is not available for this application",
            details={"app": app, "policy": policy},
        )
    return tuple(sorted(
        capability
        for capability, spec in CAPABILITIES.items()
        if spec.app == app and capability in PROFILE_CAPABILITIES[policy]
    ))


def capability_for_operation(app: str, operation: str) -> str:
    """Return the registered capability required by a semantic operation."""

    try:
        return OPERATION_CAPABILITIES[(app, operation)]
    except KeyError:
        return _operation_capability(app, operation)


def _validate_policy(policy: str) -> frozenset[str]:
    try:
        return PROFILE_CAPABILITIES[policy]
    except KeyError as exc:
        raise PolicyError(
            f"unknown policy profile: {policy}",
            details={"allowed_profiles": list(PROFILE_NAMES)},
        ) from exc


def _validate_target(capability: Capability, target: str | None) -> None:
    if capability.target == "mailbox":
        return
    if target is None or not target.strip():
        raise PolicyError(
            "capability requires a target",
            details={"capability": capability.name, "target_kind": capability.target},
        )
    if capability.target == "file" and not Path(target).is_absolute():
        raise PolicyError(
            "file capabilities require an absolute target path",
            details={"capability": capability.name, "target": target},
        )


def decide(policy: str, capability: str, *, target: str | None = None) -> dict[str, object]:
    """Return a machine-readable authorization decision without invoking an adapter."""

    grants = _validate_policy(policy)
    spec = CAPABILITIES.get(capability)
    if spec is None:
        return {
            "ok": False,
            "decision": "deny",
            "policy": policy,
            "capability": capability,
            "reason": "unknown capability",
        }
    if capability not in grants:
        return {
            "ok": False,
            "decision": "deny",
            "policy": policy,
            "capability": capability,
            "reason": "capability is not granted by this policy",
            "risk": spec.risk,
        }
    try:
        _validate_target(spec, target)
    except PolicyError as exc:
        return {
            "ok": False,
            "decision": "deny",
            "policy": policy,
            "capability": capability,
            "reason": exc.message,
            "details": exc.details,
        }
    return {
        "ok": True,
        "decision": "allow",
        "policy": policy,
        "capability": capability,
        "risk": spec.risk,
        "target": target,
    }


def require_capability(policy: str, capability: str, *, target: str | None = None) -> dict[str, object]:
    decision = decide(policy, capability, target=target)
    if not decision["ok"]:
        raise PolicyError(
            str(decision["reason"]),
            details={key: value for key, value in decision.items() if key != "ok"},
        )
    return decision


def catalog(*, policy: str | None = None, authority: Any | None = None) -> dict[str, object]:
    """Return the stable permission vocabulary for humans and agents."""

    selected = _validate_policy(policy) if policy is not None else None
    capabilities = []
    for name, spec in CAPABILITIES.items():
        entry: dict[str, object] = {
            "name": name,
            "app": spec.app,
            "description": spec.description,
            "risk": spec.risk,
            "target": spec.target,
            "profiles": [profile for profile in PROFILE_NAMES if name in PROFILE_CAPABILITIES[profile]],
        }
        if selected is not None:
            entry["profile_granted"] = name in selected
            entry["granted"] = name in selected
            if authority is not None:
                authority_granted = authority.has_capability(spec.app, name, policy=policy)
                entry["authority_granted"] = authority_granted
                entry["granted"] = entry["granted"] and authority_granted
        capabilities.append(entry)
    profiles = {
        name: sorted(grants)
        for name, grants in PROFILE_CAPABILITIES.items()
    }
    return {
        "schema": PERMISSIONS_SCHEMA,
        "profiles": profiles,
        "capabilities": capabilities,
    }


def capabilities_for_operations(app: str, operations: Iterable[str]) -> set[str]:
    return {capability_for_operation(app, operation) for operation in operations}
