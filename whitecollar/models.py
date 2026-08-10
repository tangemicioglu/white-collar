from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Literal

from .errors import ValidationError
from .mail_ops import MAIL_COM_OPERATIONS, MAIL_COM_REQUIRED_ARGS
from .slides_ops import SLIDES_COM_OPERATIONS, SLIDES_COM_REQUIRED_ARGS
from .word_ops import WORD_COM_OPERATIONS, WORD_COM_REQUIRED_ARGS

PLAN_SCHEMA = "white-collar.plan/v1"
RESULT_SCHEMA = "white-collar.result/v1"
APPS = {"word", "slides", "mail"}
POLICY_NAMES = {"read-only", "review", "edit", "send"}


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
class MailTarget:
    id: str

    @classmethod
    def from_dict(cls, raw: Any) -> "MailTarget":
        if not isinstance(raw, dict):
            raise ValidationError("mail target must be an object")
        _expect_keys(raw, required={"id"}, optional=set(), context="target")
        value = raw["id"]
        if not isinstance(value, str) or not value.strip():
            raise ValidationError("target.id must be a non-empty string")
        return cls(value)


@dataclass(frozen=True)
class WriteIntent:
    mode: Literal["none", "create", "save-as", "in-place"]
    path: str | None = None
    snapshot: str | None = None

    @classmethod
    def from_dict(cls, raw: Any, target: Target | MailTarget) -> "WriteIntent":
        if not isinstance(raw, dict):
            raise ValidationError("write must be an object")
        _expect_keys(raw, required={"mode"}, optional={"path", "snapshot"}, context="write")
        mode = raw["mode"]
        if mode == "none":
            if raw.get("path") is not None or raw.get("snapshot") is not None:
                raise ValidationError("write.none does not accept write.path or write.snapshot")
            return cls("none")
        if isinstance(target, MailTarget):
            raise ValidationError("mail plans must use write.mode 'none'")
        if mode == "create":
            if raw.get("path") is not None or raw.get("snapshot") is not None:
                raise ValidationError("write.create does not accept write.path or write.snapshot")
            return cls("create")
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
        raise ValidationError("write.mode must be 'create', 'save-as', or 'in-place'")


@dataclass(frozen=True)
class Plan:
    schema: str
    app: Literal["word", "slides", "mail"]
    target: Target | MailTarget
    policy: Literal["read-only", "review", "edit", "send"]
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
            raise ValidationError("app must be 'word', 'slides', or 'mail'")
        if raw["policy"] not in POLICY_NAMES:
            raise ValidationError("policy must be 'read-only', 'review', 'edit', or 'send'")
        target = MailTarget.from_dict(raw["target"]) if raw["app"] == "mail" else Target.from_dict(raw["target"])
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
    if not isinstance(raw.get("op"), str):
        raise ValidationError(f"{context}.op must be a string")
    operation = raw["op"]
    if operation == "replace_text":
        if app not in {"word", "slides"}:
            raise ValidationError(f"{context}.op is unsupported for {app}")
        _expect_keys(raw, required={"op", "find", "replace"}, optional={"occurrence"}, context=context)
        if not isinstance(raw["find"], str) or not raw["find"]:
            raise ValidationError(f"{context}.find must be a non-empty string")
        if not isinstance(raw["replace"], str):
            raise ValidationError(f"{context}.replace must be a string")
        occurrence = raw.get("occurrence", "all")
        if occurrence not in {"all", "first"}:
            raise ValidationError(f"{context}.occurrence must be 'all' or 'first'")
        return {"op": operation, "find": raw["find"], "replace": raw["replace"], "occurrence": occurrence}
    if app == "word" and operation in WORD_COM_OPERATIONS:
        _expect_keys(raw, required={"op"}, optional={"args"}, context=context)
        args = raw.get("args", {})
        if not isinstance(args, dict):
            raise ValidationError(f"{context}.args must be an object")
        missing = WORD_COM_REQUIRED_ARGS.get(operation, set()) - set(args)
        if missing:
            raise ValidationError(f"{context}.args is missing required field(s): {', '.join(sorted(missing))}")
        if operation == "word_live_add_table" and not ("columns" in args or "cols" in args):
            raise ValidationError(f"{context}.args requires columns or cols")
        if operation == "word_live_find_text" and not ("find_text" in args or "search_text" in args):
            raise ValidationError(f"{context}.args is missing required field: search_text or find_text")
        if operation in {"word_live_delete_text", "word_live_format_text", "word_live_get_paragraph_format", "word_live_set_paragraph_spacing"}:
            has_range = {"start", "end"}.issubset(args) or "start_paragraph" in args or "paragraph_index" in args
            if not has_range:
                raise ValidationError(f"{context}.args must identify a character or paragraph range")
        if operation == "word_live_add_comment" and not ("comment_text" in args or "text" in args):
            raise ValidationError(f"{context}.args requires text or comment_text")
        if operation == "word_live_add_comment" and not (
            {"start", "end"}.issubset(args) or "start_paragraph" in args or "paragraph_index" in args or "target_text" in args
        ):
            raise ValidationError(f"{context}.args must identify a comment target range")
        if operation == "word_live_reply_to_comment" and not ("reply_text" in args or "text" in args):
            raise ValidationError(f"{context}.args requires text or reply_text")
        return {"op": operation, "args": args}
    if app == "slides" and operation in SLIDES_COM_OPERATIONS:
        _expect_keys(raw, required={"op"}, optional={"args"}, context=context)
        args = raw.get("args", {})
        if not isinstance(args, dict):
            raise ValidationError(f"{context}.args must be an object")
        missing = SLIDES_COM_REQUIRED_ARGS.get(operation, set()) - set(args)
        if missing:
            raise ValidationError(f"{context}.args is missing required field(s): {', '.join(sorted(missing))}")
        return {"op": operation, "args": args}
    if app == "mail" and operation in MAIL_COM_OPERATIONS:
        _expect_keys(raw, required={"op"}, optional={"args"}, context=context)
        args = raw.get("args", {})
        if not isinstance(args, dict):
            raise ValidationError(f"{context}.args must be an object")
        missing = MAIL_COM_REQUIRED_ARGS.get(operation, set()) - set(args)
        if missing:
            raise ValidationError(f"{context}.args is missing required field(s): {', '.join(sorted(missing))}")
        if operation == "mail_live_create_draft":
            _expect_keys(
                args,
                required={"to", "subject", "body"},
                optional={"cc", "bcc"},
                context=f"{context}.args",
            )
            for field in ("to", "subject", "body"):
                if not isinstance(args[field], str):
                    raise ValidationError(f"{context}.args.{field} must be a string")
            if not args["to"].strip():
                raise ValidationError(f"{context}.args.to must be a non-empty string")
            if not args["subject"].strip():
                raise ValidationError(f"{context}.args.subject must be a non-empty string")
            for field in ("cc", "bcc"):
                if field in args and not isinstance(args[field], str):
                    raise ValidationError(f"{context}.args.{field} must be a string")
        return {"op": operation, "args": args}
    raise ValidationError(f"{context}.op is unsupported for {app}")


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
    }
    if dry_run:
        envelope["dry_run"] = True
    if target is not None:
        envelope["target"] = target
    if data is not None:
        envelope["data"] = data
    if changes is not None:
        envelope["changes"] = changes
    if error is not None:
        envelope["error"] = error
    return envelope
