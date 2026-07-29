from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Literal

from .errors import ValidationError

PLAN_SCHEMA = "white-collar.plan/v1"
RESULT_SCHEMA = "white-collar.result/v1"
APPS = {"word", "slides"}
POLICY_NAMES = {"read-only", "review", "edit"}


def _expect_keys(raw: dict[str, Any], *, required: set[str], optional: set[str], context: str) -> None:
    missing = required - set(raw)
    unknown = set(raw) - required - optional
    if missing:
        raise ValidationError(f"{context} is missing required field(s): {', '.join(sorted(missing))}")
    if unknown:
        raise ValidationError(f"{context} has unknown field(s): {', '.join(sorted(unknown))}")


def _absolute_path(value: Any, field: str) -> str:
    if not isinstance(value, str) or not value:
        raise ValidationError(f"{field} must be a non-empty string")
    if not Path(value).is_absolute():
        raise ValidationError(f"{field} must be an absolute path")
    return str(Path(value))


@dataclass(frozen=True)
class Target:
    path: str
    expected_sha256: str | None = None

    @classmethod
    def from_dict(cls, raw: Any) -> "Target":
        if not isinstance(raw, dict):
            raise ValidationError("target must be an object")
        _expect_keys(raw, required={"path"}, optional={"expected_sha256"}, context="target")
        digest = raw.get("expected_sha256")
        if digest is not None and (
            not isinstance(digest, str)
            or len(digest) != 64
            or any(character not in "0123456789abcdefABCDEF" for character in digest)
        ):
            raise ValidationError("target.expected_sha256 must be a 64-character hexadecimal SHA-256 digest")
        return cls(_absolute_path(raw["path"], "target.path"), digest)


@dataclass(frozen=True)
class WriteIntent:
    mode: Literal["save-as", "in-place"]
    path: str | None = None
    snapshot: str | None = None

    @classmethod
    def from_dict(cls, raw: Any, target: Target) -> "WriteIntent":
        if not isinstance(raw, dict):
            raise ValidationError("write must be an object")
        _expect_keys(raw, required={"mode"}, optional={"path", "snapshot"}, context="write")
        mode = raw["mode"]
        if mode == "save-as":
            path = _absolute_path(raw.get("path"), "write.path")
            if Path(path) == Path(target.path):
                raise ValidationError("save-as path must differ from the target")
            if raw.get("snapshot") is not None:
                raise ValidationError("save-as does not accept write.snapshot")
            return cls("save-as", path=path)
        if mode == "in-place":
            snapshot = _absolute_path(raw.get("snapshot"), "write.snapshot")
            if Path(snapshot) == Path(target.path):
                raise ValidationError("snapshot path must differ from the target")
            if raw.get("path") is not None:
                raise ValidationError("in-place does not accept write.path")
            return cls("in-place", snapshot=snapshot)
        raise ValidationError("write.mode must be 'save-as' or 'in-place'")


@dataclass(frozen=True)
class Plan:
    schema: str
    app: Literal["word", "slides"]
    target: Target
    policy: Literal["read-only", "review", "edit"]
    operations: tuple[dict[str, Any], ...]
    write: WriteIntent

    @classmethod
    def from_dict(cls, raw: Any) -> "Plan":
        if not isinstance(raw, dict):
            raise ValidationError("plan must be a JSON object")
        _expect_keys(
            raw,
            required={"schema", "app", "target", "policy", "operations", "write"},
            optional=set(),
            context="plan",
        )
        if raw["schema"] != PLAN_SCHEMA:
            raise ValidationError(f"schema must be {PLAN_SCHEMA!r}")
        if raw["app"] not in APPS:
            raise ValidationError("app must be 'word' or 'slides'")
        if raw["policy"] not in POLICY_NAMES:
            raise ValidationError("policy must be 'read-only', 'review', or 'edit'")
        target = Target.from_dict(raw["target"])
        write = WriteIntent.from_dict(raw["write"], target)
        operations = raw["operations"]
        if not isinstance(operations, list) or not operations:
            raise ValidationError("operations must be a non-empty array")
        normalized = tuple(_validate_operation(raw["app"], operation, index) for index, operation in enumerate(operations))
        return cls(PLAN_SCHEMA, raw["app"], target, raw["policy"], normalized, write)

    def to_dict(self) -> dict[str, Any]:
        value = asdict(self)
        value["operations"] = list(self.operations)
        return value


def _validate_operation(app: str, raw: Any, index: int) -> dict[str, Any]:
    context = f"operations[{index}]"
    if not isinstance(raw, dict):
        raise ValidationError(f"{context} must be an object")
    _expect_keys(raw, required={"op", "find", "replace"}, optional={"occurrence"}, context=context)
    if raw["op"] != "replace_text":
        raise ValidationError(f"{context}.op is unsupported for {app}")
    if not isinstance(raw["find"], str) or not raw["find"]:
        raise ValidationError(f"{context}.find must be a non-empty string")
    if not isinstance(raw["replace"], str):
        raise ValidationError(f"{context}.replace must be a string")
    occurrence = raw.get("occurrence", "all")
    if occurrence not in {"all", "first"}:
        raise ValidationError(f"{context}.occurrence must be 'all' or 'first'")
    return {"op": "replace_text", "find": raw["find"], "replace": raw["replace"], "occurrence": occurrence}


def result(
    *,
    ok: bool,
    command: str,
    policy: str,
    dry_run: bool,
    target: str | None = None,
    data: dict[str, Any] | list[Any] | None = None,
    changes: list[dict[str, Any]] | None = None,
    error: dict[str, Any] | None = None,
) -> dict[str, Any]:
    envelope: dict[str, Any] = {
        "schema": RESULT_SCHEMA,
        "ok": ok,
        "command": command,
        "policy": policy,
        "dry_run": dry_run,
    }
    if target is not None:
        envelope["target"] = target
    if data is not None:
        envelope["data"] = data
    if changes is not None:
        envelope["changes"] = changes
    if error is not None:
        envelope["error"] = error
    return envelope
