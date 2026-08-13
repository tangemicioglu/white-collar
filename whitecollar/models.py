from __future__ import annotations

import ntpath
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
    # Plans describe files consumed by Windows Office even when they are
    # validated on a POSIX CI runner.  pathlib.Path follows the host OS, so
    # also recognize Windows drive and UNC paths explicitly.
    if not Path(value).is_absolute() and not ntpath.isabs(value):
        raise ValidationError(f"{field} must be an absolute path")
    # Normalize Windows paths with Windows semantics even when validation is
    # running on a POSIX CI runner.  This keeps the representation passed to
    # COM stable across platforms.
    if Path(value).is_absolute():
        return str(Path(value))
    return ntpath.normpath(value) if ntpath.isabs(value) else str(Path(value))


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
    display: dict[str, Any] | None = None

    @classmethod
    def from_dict(cls, raw: Any) -> "Plan":
        if not isinstance(raw, dict):
            raise ValidationError("plan must be a JSON object")
        _expect_keys(
            raw,
            required={"schema", "app", "target", "policy", "operations", "write"},
            optional={"display"},
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
        display = raw.get("display")
        if display is not None:
            if not isinstance(display, dict):
                raise ValidationError("display must be an object")
            _expect_keys(
                display,
                required=set(),
                optional={"pause_after_operation", "keep_live_as_output"},
                context="display",
            )
            pause = display.get("pause_after_operation", 0)
            if isinstance(pause, bool) or not isinstance(pause, (int, float)) or pause < 0 or pause > 10:
                raise ValidationError("display.pause_after_operation must be between 0 and 10 seconds")
            keep_live = display.get("keep_live_as_output", False)
            if not isinstance(keep_live, bool):
                raise ValidationError("display.keep_live_as_output must be a boolean")
            display = {"pause_after_operation": float(pause), "keep_live_as_output": keep_live}
        return cls(PLAN_SCHEMA, raw["app"], target, raw["policy"], normalized, write, display)

    def to_dict(self) -> dict[str, Any]:
        value = asdict(self)
        value["operations"] = list(self.operations)
        if value.get("display") is None:
            value.pop("display", None)
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
        if operation == "word_live_remove_watermark":
            _expect_keys(
                args,
                required=set(),
                optional={"text", "section_index", "position"},
                context=f"{context}.args",
            )
            if "text" in args and (not isinstance(args["text"], str) or not args["text"].strip()):
                raise ValidationError(f"{context}.args.text must be a non-empty string")
            if "section_index" in args and (
                isinstance(args["section_index"], bool)
                or not isinstance(args["section_index"], int)
                or args["section_index"] < 1
            ):
                raise ValidationError(f"{context}.args.section_index must be a positive integer")
            if args.get("position", "both") not in {"header", "footer", "both"}:
                raise ValidationError(f"{context}.args.position must be 'header', 'footer', or 'both'")
            return {"op": operation, "args": args}
        if operation in {
            "word_live_list_styles", "word_live_list_hyperlinks", "word_live_list_notes",
            "word_live_list_content_controls", "word_live_get_protection", "word_live_update_fields",
            "word_live_export_pdf",
        }:
            _expect_keys(args, required=set(), optional=set(), context=f"{context}.args")
            return {"op": operation, "args": args}
        if operation == "word_live_apply_style":
            _expect_keys(
                args,
                required={"style_name"},
                optional=_WORD_RANGE_ARGS,
                context=f"{context}.args",
            )
            _require_non_empty_string(args, "style_name", f"{context}.args")
            return {"op": operation, "args": args}
        if operation == "word_live_add_hyperlink":
            _expect_keys(
                args,
                required={"url"},
                optional=_WORD_RANGE_ARGS | {"sub_address", "display_text"},
                context=f"{context}.args",
            )
            _require_non_empty_string(args, "url", f"{context}.args")
            return {"op": operation, "args": args}
        if operation == "word_live_remove_hyperlink":
            _expect_keys(
                args,
                required=set(),
                optional=_WORD_RANGE_ARGS | {"hyperlink_index"},
                context=f"{context}.args",
            )
            if "hyperlink_index" not in args and not _has_word_range(args):
                raise ValidationError(f"{context}.args must identify a hyperlink or range")
            return {"op": operation, "args": args}
        if operation == "word_live_add_note":
            _expect_keys(
                args,
                required={"text"},
                optional=_WORD_RANGE_ARGS | {"note_type"},
                context=f"{context}.args",
            )
            _require_non_empty_string(args, "text", f"{context}.args")
            if args.get("note_type", "footnote") not in {"footnote", "endnote"}:
                raise ValidationError(f"{context}.args.note_type must be 'footnote' or 'endnote'")
            return {"op": operation, "args": args}
        if operation == "word_live_insert_toc":
            _expect_keys(
                args,
                required=set(),
                optional={
                    "position", "bookmark", "use_heading_styles", "upper_heading_level",
                    "lower_heading_level", "right_align_page_numbers", "include_page_numbers",
                },
                context=f"{context}.args",
            )
            if args.get("position", "start") not in {"start", "end", "cursor"} and not isinstance(args.get("position"), int):
                raise ValidationError(f"{context}.args.position must be start, end, cursor, or a character offset")
            return {"op": operation, "args": args}
        if operation == "word_live_set_content_control":
            _expect_keys(
                args,
                required={"title", "value"},
                optional=_WORD_RANGE_ARGS | {"tag", "create_if_missing"},
                context=f"{context}.args",
            )
            _require_non_empty_string(args, "title", f"{context}.args")
            if not isinstance(args["value"], str):
                raise ValidationError(f"{context}.args.value must be a string")
            return {"op": operation, "args": args}
        if operation == "word_live_remove_header_footer":
            _expect_keys(args, required=set(), optional={"position", "section_index"}, context=f"{context}.args")
            if args.get("position", "both") not in {"header", "footer", "both"}:
                raise ValidationError(f"{context}.args.position must be 'header', 'footer', or 'both'")
            return {"op": operation, "args": args}
        if operation == "word_live_set_protection":
            _expect_keys(args, required={"protection_type"}, optional={"password"}, context=f"{context}.args")
            if args["protection_type"] not in {"none", "tracked_changes", "comments", "forms", "read_only"}:
                raise ValidationError(f"{context}.args.protection_type is invalid")
            return {"op": operation, "args": args}
        if operation in {"word_live_compare_documents", "word_live_merge_document"}:
            _expect_keys(args, required={"source_path"}, optional=set(), context=f"{context}.args")
            _absolute_path(args["source_path"], f"{context}.args.source_path")
            return {"op": operation, "args": args}
        missing = WORD_COM_REQUIRED_ARGS.get(operation, set()) - set(args)
        if missing:
            raise ValidationError(f"{context}.args is missing required field(s): {', '.join(sorted(missing))}")
        if operation == "word_live_add_table" and not ("columns" in args or "cols" in args):
            raise ValidationError(f"{context}.args requires columns or cols")
        if operation == "word_live_find_text" and not ("find_text" in args or "search_text" in args):
            raise ValidationError(f"{context}.args is missing required field: search_text or find_text")
        if operation in {"word_live_delete_text", "word_live_format_text", "word_live_get_paragraph_format", "word_live_set_paragraph_spacing"}:
            if not _has_word_range(args):
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
        _validate_slides_operation(operation, args, context)
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


_WORD_RANGE_ARGS = {
    "start", "end", "start_paragraph", "end_paragraph", "paragraph_index",
    "bookmark", "target_text",
}


def _require_non_empty_string(args: dict[str, Any], name: str, context: str) -> None:
    value = args.get(name)
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{context}.{name} must be a non-empty string")


def _has_word_range(args: dict[str, Any]) -> bool:
    return (
        {"start", "end"}.issubset(args)
        or "start_paragraph" in args
        or "paragraph_index" in args
        or "bookmark" in args
        or "target_text" in args
    )


def _validate_slides_operation(operation: str, args: dict[str, Any], context: str) -> None:
    shape_selector = {"shape_name", "shape_index", "slide_index"}
    if operation in {"slides_live_get_masters", "slides_live_get_layouts", "slides_live_get_placeholders", "slides_live_get_notes", "slides_live_get_sections", "slides_live_get_media"}:
        _expect_keys(
            args,
            required=set(),
            optional={"slide_index", "master"},
            context=f"{context}.args",
        )
    elif operation == "slides_live_apply_template":
        _expect_keys(args, required={"source_path"}, optional=set(), context=f"{context}.args")
        _absolute_path(args["source_path"], f"{context}.args.source_path")
    elif operation in {"slides_live_save_template", "slides_live_export_pdf"}:
        _expect_keys(args, required=set(), optional=set(), context=f"{context}.args")
    elif operation == "slides_live_set_layout":
        _expect_keys(args, required={"layout"}, optional={"slide_index", "master"}, context=f"{context}.args")
    elif operation == "slides_live_group":
        _expect_keys(args, required=set(), optional={"shape_names", "shape_indices", "slide_index"}, context=f"{context}.args")
        _validate_shape_list(args, context)
    elif operation == "slides_live_ungroup":
        _expect_keys(args, required=set(), optional=shape_selector, context=f"{context}.args")
        _require_shape_selector(args, context)
    elif operation in {"slides_live_align", "slides_live_distribute"}:
        required = "alignment" if operation == "slides_live_align" else "direction"
        _expect_keys(
            args,
            required={required},
            optional={"shape_names", "shape_indices", "slide_index", "relative_to_slide"},
            context=f"{context}.args",
        )
        # Names and indexes are alternative bounded selectors.  Keep the
        # public plan vocabulary independent of PowerPoint's ShapeRange API.
        if "shape_names" not in args and "shape_indices" not in args:
            raise ValidationError(f"{context}.args requires shape_names or shape_indices")
        _validate_shape_list(args, context)
    elif operation == "slides_live_z_order":
        _expect_keys(args, required={"command"}, optional=shape_selector - {"slide_index"} | {"slide_index"}, context=f"{context}.args")
        _require_shape_selector(args, context)
    elif operation == "slides_live_crop_image":
        _expect_keys(
            args,
            required=set(),
            optional=shape_selector | {"left", "top", "right", "bottom"},
            context=f"{context}.args",
        )
        _require_shape_selector(args, context)
        if not any(key in args for key in ("left", "top", "right", "bottom")):
            raise ValidationError(f"{context}.args requires at least one crop value")
    elif operation == "slides_live_rotate_shape":
        _expect_keys(args, required={"degrees"}, optional=shape_selector, context=f"{context}.args")
        _require_shape_selector(args, context)
    elif operation == "slides_live_add_section":
        _expect_keys(args, required={"name"}, optional={"slide_index"}, context=f"{context}.args")
    elif operation == "slides_live_delete_section":
        _expect_keys(args, required=set(), optional={"section_index"}, context=f"{context}.args")
    elif operation == "slides_live_set_slide_visibility":
        _expect_keys(args, required={"visible"}, optional={"slide_index"}, context=f"{context}.args")
    elif operation == "slides_live_set_slide_numbers":
        _expect_keys(args, required={"visible"}, optional={"slide_index"}, context=f"{context}.args")
    elif operation == "slides_live_add_table":
        _expect_keys(
            args,
            required={"rows", "columns"},
            optional={"slide_index", "name", "data", "left", "top", "width", "height"},
            context=f"{context}.args",
        )
    elif operation == "slides_live_set_table_cell":
        _expect_keys(args, required={"row", "column", "text"}, optional=shape_selector, context=f"{context}.args")
        _require_shape_selector(args, context, include_slide=False)
    elif operation == "slides_live_add_chart":
        _expect_keys(
            args,
            required={"chart_type"},
            optional={"slide_index", "name", "title", "data", "left", "top", "width", "height"},
            context=f"{context}.args",
        )
        if "data" in args and (
            not isinstance(args["data"], list)
            or not args["data"]
            or not all(isinstance(row, list) and row for row in args["data"])
        ):
            raise ValidationError(f"{context}.args.data must be a non-empty array of non-empty rows")
    elif operation == "slides_live_add_smartart":
        _expect_keys(args, required=set(), optional={"slide_index", "name", "layout", "left", "top", "width", "height", "nodes"}, context=f"{context}.args")
        if "nodes" in args and (
            not isinstance(args["nodes"], list)
            or not args["nodes"]
            or not all(isinstance(item, str) for item in args["nodes"])
        ):
            raise ValidationError(f"{context}.args.nodes must be a non-empty array of strings")
    elif operation == "slides_live_add_media":
        _expect_keys(args, required={"media_path"}, optional={"slide_index", "name", "left", "top", "width", "height"}, context=f"{context}.args")
        _absolute_path(args["media_path"], f"{context}.args.media_path")
    elif operation == "slides_live_set_hyperlink":
        _expect_keys(args, required={"url"}, optional=shape_selector | {"sub_address"}, context=f"{context}.args")
        _require_shape_selector(args, context)
    elif operation == "slides_live_set_alt_text":
        _expect_keys(args, required={"text"}, optional=shape_selector, context=f"{context}.args")
        _require_shape_selector(args, context)
    elif operation == "slides_live_set_transition":
        _expect_keys(args, required={"effect"}, optional={"slide_index", "advance_on_click", "advance_seconds"}, context=f"{context}.args")
    elif operation == "slides_live_add_animation":
        _expect_keys(args, required={"effect"}, optional=shape_selector | {"trigger"}, context=f"{context}.args")
        _require_shape_selector(args, context)


def _validate_shape_list(args: dict[str, Any], context: str) -> None:
    names = args.get("shape_names")
    indices = args.get("shape_indices")
    if names is None and indices is None:
        raise ValidationError(f"{context}.args requires shape_names or shape_indices")
    if names is not None and (
        not isinstance(names, list)
        or len(names) < 1
        or not all(isinstance(item, str) and item for item in names)
    ):
        raise ValidationError(f"{context}.args.shape_names must be a non-empty array of strings")
    if indices is not None and (
        not isinstance(indices, list)
        or len(indices) < 1
        or not all(isinstance(item, int) and not isinstance(item, bool) and item > 0 for item in indices)
    ):
        raise ValidationError(f"{context}.args.shape_indices must be an array of positive integers")


def _require_shape_selector(args: dict[str, Any], context: str, *, include_slide: bool = True) -> None:
    allowed = {"shape_name", "shape_index"}
    if include_slide:
        allowed.add("slide_index")
    if not any(key in args for key in allowed if key != "slide_index"):
        raise ValidationError(f"{context}.args must identify a shape")


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
