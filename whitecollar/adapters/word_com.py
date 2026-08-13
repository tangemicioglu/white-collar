from __future__ import annotations

import contextlib
import datetime as dt
import hashlib
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any, Callable, Iterator

from ..errors import BackendUnavailableError, ValidationError
from ..models import Plan
from ..office_capture import capture_window
from ..word_ops import WORD_COM_MUTATING_OPERATIONS, WORD_COM_OPERATIONS

WORD_COM_SELF_WRITING_OPERATIONS = {"word_live_export_pdf", "word_live_compare_documents"}


class Win32WordComAdapter:
    """Finite semantic adapter for the Word live operation vocabulary.

    The adapter owns the COM details. Plans can name only the operations in
    ``word_ops.py`` and their documented argument objects; they cannot invoke
    arbitrary COM methods.
    """

    def __init__(
        self,
        *,
        app_factory: Callable[[], Any] | None = None,
        screenshotter: Callable[[int, Path], None] | None = None,
    ):
        self._app_factory = app_factory or _default_word_app
        self._screenshotter = screenshotter or capture_window
        self._snapshots: dict[str, list[dict[str, Any]]] = {}
        self._history: list[dict[str, Any]] = []
        self._pending_comment_package_edits: dict[str, list[dict[str, Any]]] = {}

    def inspect(self, target: Path, *, render_dir: Path | None = None) -> dict[str, Any]:
        app = self._get_app()
        doc, opened_here = _find_or_open_for_inspect(app, target)
        try:
            data = self._get_info(doc) | {"backend": "word-com", "text": self._get_text(doc)["paragraphs"]}
            if render_dir is not None:
                data["renders"] = _render_document(doc, render_dir)
            return data
        finally:
            if opened_here:
                doc.Close(SaveChanges=False)

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
        if any(operation["op"] == "word_live_create_document" for operation in plan.operations):
            return self._word_live_create_document(plan, dry_run=dry_run)
        if dry_run and all(operation["op"] in WORD_COM_MUTATING_OPERATIONS or operation["op"] == "replace_text" for operation in plan.operations):
            return {
                "backend": "word-com",
                "written": False,
                "operations": [
                    {"op": operation["op"], "dry_run": True, "args": _operation_args(operation)}
                    for operation in plan.operations
                ],
            }
        app = self._get_app()
        doc = None
        if any(operation["op"] != "word_live_list_open" for operation in plan.operations):
            doc = _find_document(app, plan.target.path)
        if not dry_run and doc is not None and plan.write.mode != "none":
            self._preflight_target(plan)
            self._prepare_write(doc, plan)
        operations = []
        for operation in plan.operations:
            name = operation["op"]
            args = _operation_args(operation)
            if name in WORD_COM_SELF_WRITING_OPERATIONS and plan.write.path:
                args.setdefault("output_path", plan.write.path)
            if name not in WORD_COM_OPERATIONS and name != "replace_text":
                raise ValidationError(f"unsupported Word COM operation: {name}")
            if dry_run and name in WORD_COM_MUTATING_OPERATIONS:
                operations.append({"op": name, "dry_run": True, "args": args})
                continue
            if name == "word_live_list_open":
                operations.append(self._list_open(app))
                continue
            if doc is None:
                doc = _find_document(app, plan.target.path)
            dispatch_name = "word_live_replace_text" if name == "replace_text" else name
            method = getattr(self, f"_{dispatch_name}", None)
            if method is None:
                raise ValidationError(f"Word operation is registered but not implemented: {name}")
            with self._undo(app, name) if name in WORD_COM_MUTATING_OPERATIONS else contextlib.nullcontext():
                operations.append(method(app, doc, args))
            self._history.append({"operation": name, "document": doc.Name, "at": dt.datetime.now(dt.timezone.utc).isoformat()})
        if not dry_run and doc is not None and plan.write.mode != "none" and not any(
            operation["op"] in WORD_COM_SELF_WRITING_OPERATIONS for operation in plan.operations
        ):
            self._commit_write(app, doc, plan)
        return {"backend": "word-com", "written": not dry_run, "operations": operations}

    def _word_live_create_document(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
        if len(plan.operations) != 1 or plan.operations[0]["op"] != "word_live_create_document":
            raise ValidationError("word_live_create_document must be a standalone plan")
        target = Path(plan.target.path)
        if target.suffix.lower() != ".docx":
            raise ValidationError("Word creation output must use the .docx extension", details={"target": str(target)})
        if target.exists():
            raise ValidationError("creation output already exists", details={"target": str(target)})
        if dry_run:
            return {
                "backend": "word-com",
                "written": False,
                "operations": [{"op": "word_live_create_document", "dry_run": True, "args": {}}],
            }
        app = self._get_app()
        document = None
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            document = app.Documents.Add()
            save_as2 = getattr(document, "SaveAs2", None)
            if callable(save_as2):
                save_as2(FileName=str(target), FileFormat=12)
            else:
                document.SaveAs(FileName=str(target), FileFormat=12)
            if not target.is_file():
                raise ValidationError("Word did not create the requested document", details={"target": str(target)})
            operation = {"op": "word_live_create_document", "created": True, "path": str(target)}
            return {"backend": "word-com", "written": True, "operations": [operation]}
        except ValidationError:
            raise
        except Exception as exc:
            raise ValidationError(
                "Word could not create the requested document",
                details={"target": str(target), "reason": str(exc)},
            ) from exc
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:
                    pass

    def _preflight_target(self, plan: Plan) -> None:
        target = Path(plan.target.path)
        if plan.target.expected_sha256 and target.is_file():
            actual = _sha256(target)
            if actual.lower() != plan.target.expected_sha256.lower():
                raise ValidationError(
                    "target SHA-256 does not match plan",
                    details={"expected": plan.target.expected_sha256, "actual": actual},
                )
        if plan.write.mode == "save-as":
            output = Path(plan.write.path or "")
            if output.exists():
                raise ValidationError("save-as path already exists", details={"path": str(output)})

    def _prepare_write(self, doc: Any, plan: Plan) -> None:
        if plan.write.mode != "in-place":
            return
        snapshot = Path(plan.write.snapshot or "")
        if snapshot.exists():
            raise ValidationError("snapshot already exists", details={"path": str(snapshot)})
        snapshot.parent.mkdir(parents=True, exist_ok=True)
        save_copy_as = getattr(doc, "SaveCopyAs", None)
        if callable(save_copy_as):
            try:
                # Some Word builds expose SaveCopyAs but reject every call with
                # 0x800a1704. Keep the native path when it works (and for
                # mockable adapters), then use the on-disk document as the
                # faithful fallback for those builds.
                save_copy_as(str(snapshot))
                return
            except Exception:
                pass
        source = Path(str(_safe_value(doc, "FullName", "")))
        if not source.is_file():
            raise ValidationError(
                "Word COM could not create a snapshot and the open document has no on-disk file",
                details={"source": str(source), "snapshot": str(snapshot)},
            )
        if _safe_value(doc, "Saved", True) is False:
            doc.Save()
        shutil.copy2(source, snapshot)

    def _commit_write(self, app: Any, doc: Any, plan: Plan) -> None:
        if plan.write.mode == "in-place":
            doc.Save()
            self._commit_pending_comment_edits(app, doc, Path(str(_safe_value(doc, "FullName", ""))), reopen=True)
            return
        output = Path(plan.write.path or "")
        if output.exists():
            raise ValidationError("save-as path already exists", details={"path": str(output)})
        output.parent.mkdir(parents=True, exist_ok=True)
        save_copy_as = getattr(doc, "SaveCopyAs", None)
        if callable(save_copy_as):
            try:
                save_copy_as(str(output))
            except Exception:
                # SaveCopyAs is broken in some current Word installations.
                # SaveAs2 below is reliable, but changes the live document's
                # name, so restore the original open target after writing.
                source = Path(str(_safe_value(doc, "FullName", "")))
                doc.SaveAs2(FileName=str(output), FileFormat=_file_format(output))
                if source.resolve() != output.resolve():
                    doc.Close(SaveChanges=False)
                    doc = app.Documents.Open(FileName=str(source), ReadOnly=False, AddToRecentFiles=False)
        else:
            source = Path(str(_safe_value(doc, "FullName", "")))
            doc.SaveAs2(FileName=str(output), FileFormat=_file_format(output))
            if source.resolve() != output.resolve():
                doc.Close(SaveChanges=False)
                doc = app.Documents.Open(FileName=str(source), ReadOnly=False, AddToRecentFiles=False)
        self._commit_pending_comment_edits(app, doc, output, reopen=False)

    def _get_app(self) -> Any:
        try:
            return self._app_factory()
        except ImportError as exc:
            raise BackendUnavailableError("word-com") from exc
        except OSError as exc:
            raise BackendUnavailableError("word-com") from exc

    def _undo(self, app: Any, name: str):
        return _undo_record(app, f"white-collar: {name}")

    def _list_open(self, app: Any) -> dict[str, Any]:
        documents = []
        for doc in _iter_collection(getattr(app, "Documents", [])):
            documents.append(
                {
                    "name": _safe_value(doc, "Name"),
                    "path": _safe_value(doc, "FullName"),
                    "saved": _safe_value(doc, "Saved"),
                    "track_revisions": _safe_value(doc, "TrackRevisions"),
                }
            )
        return {"op": "word_live_list_open", "documents": documents}

    def _word_live_get_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        paragraphs = []
        for index, paragraph in enumerate(_iter_collection(doc.Paragraphs), start=1):
            text = str(_safe_value(paragraph.Range, "Text", "")).rstrip("\r\a")
            paragraphs.append({"index": index, "text": text, "style": _style_name(paragraph)})
        return {"op": "word_live_get_text", "paragraphs": paragraphs, "total_paragraphs": len(paragraphs)}

    def _word_live_list_styles(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        styles = []
        for style in _iter_collection(doc.Styles):
            styles.append({
                "name": str(_safe_value(style, "NameLocal", _safe_value(style, "Name", ""))),
                "type": _safe_value(style, "Type"),
                "built_in": bool(_safe_value(style, "BuiltIn", False)),
            })
        return {"op": "word_live_list_styles", "styles": styles, "count": len(styles)}

    def _word_live_list_hyperlinks(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        links = []
        for index, link in enumerate(_iter_collection(doc.Hyperlinks), start=1):
            anchor = _safe_value(link, "Range")
            links.append({
                "index": index,
                "text": str(_safe_value(anchor, "Text", "")),
                "address": str(_safe_value(link, "Address", "") or ""),
                "sub_address": str(_safe_value(link, "SubAddress", "") or ""),
            })
        return {"op": "word_live_list_hyperlinks", "hyperlinks": links, "count": len(links)}

    def _word_live_list_notes(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        notes = []
        for note_type, collection_name in (("footnote", "Footnotes"), ("endnote", "Endnotes")):
            for index, note in enumerate(_iter_collection(getattr(doc, collection_name)), start=1):
                notes.append({
                    "type": note_type,
                    "index": index,
                    "text": str(_safe_value(_safe_value(note, "Range"), "Text", "")).rstrip("\r\a"),
                    "reference": str(_safe_value(_safe_value(note, "Reference"), "Text", "")),
                })
        return {"op": "word_live_list_notes", "notes": notes, "count": len(notes)}

    def _word_live_list_content_controls(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        controls = []
        collection = doc.ContentControls
        for index in range(1, _count(collection) + 1):
            control = collection(index)
            controls.append({
                "index": index,
                "id": _safe_value(control, "ID"),
                "title": str(_safe_value(control, "Title", "")),
                "tag": str(_safe_value(control, "Tag", "")),
                "type": _safe_value(control, "Type"),
                "text": str(_safe_value(_safe_value(control, "Range"), "Text", "")).rstrip("\r\a"),
                "locked": bool(_safe_value(control, "LockContents", False)),
            })
        return {"op": "word_live_list_content_controls", "content_controls": controls, "count": len(controls)}

    def _word_live_get_protection(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        protection_type = _safe_value(doc, "ProtectionType", -1)
        return {"op": "word_live_get_protection", "protection_type": protection_type, "protected": protection_type != -1}

    def _get_text(self, doc: Any) -> dict[str, Any]:
        return self._word_live_get_text(None, doc, {})

    def _get_info(self, doc: Any) -> dict[str, Any]:
        value = self._word_live_get_info(None, doc, {})
        value.pop("op", None)
        return value

    def _word_live_get_info(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        try:
            words = int(doc.ComputeStatistics(0))
        except Exception:
            words = None
        return {
            "op": "word_live_get_info",
            "name": _safe_value(doc, "Name"),
            "path": _safe_value(doc, "FullName"),
            "paragraphs": _count(doc.Paragraphs),
            "sections": _count(doc.Sections),
            "words": words,
            "track_revisions": bool(_safe_value(doc, "TrackRevisions", False)),
            "saved": bool(_safe_value(doc, "Saved", False)),
        }

    def _word_live_insert_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        text = _required_string(args, "text")
        text = _word_text(text)
        position = args.get("position", "end")
        bookmark = args.get("bookmark")
        rng = _insert_range(doc, position, bookmark)
        tracking = bool(args["track_changes"]) if "track_changes" in args else None
        with _tracking(doc, tracking):
            chunks = [text[i : i + 30000] for i in range(0, max(len(text), 1), 30000)]
            if position == "start":
                for chunk in reversed(chunks):
                    rng.InsertBefore(chunk)
            else:
                for chunk in chunks:
                    rng.InsertAfter(chunk)
                    _collapse_end(rng)
        return {"op": "word_live_insert_text", "text_length": len(text), "position": position, "chunks_used": len(chunks)}

    def _word_live_delete_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        rng = _resolve_range(doc, args)
        tracking = bool(args["track_changes"]) if "track_changes" in args else None
        with _tracking(doc, tracking):
            deleted = str(_safe_value(rng, "Text", ""))
            rng.Delete()
        return {"op": "word_live_delete_text", "characters": len(deleted)}

    def _word_live_replace_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        find_text = _required_string(args, "find_text")
        replacement = str(args.get("replace_text", ""))
        replace_all = bool(args.get("replace_all", True))
        matches: list[tuple[int, int]] = []
        cursor = int(doc.Content.Start)
        content_end = int(doc.Content.End)
        while cursor < content_end:
            rng = doc.Range(cursor, content_end)
            finder = _configure_find(rng.Find, find_text, args)
            if not finder.Execute():
                break
            start, end = int(rng.Start), int(rng.End)
            if end <= cursor:
                break
            matches.append((start, end))
            if not replace_all:
                break
            cursor = end

        tracking = bool(args["track_changes"]) if "track_changes" in args else None
        with _tracking(doc, tracking):
            # Replace from the end so earlier character offsets remain valid.
            # Assigning Range.Text is reliable across current Word builds and
            # preserves real tracked revisions when tracking is enabled.
            for start, end in reversed(matches):
                doc.Range(start, end).Text = replacement
        return {"op": "word_live_replace_text", "find": find_text, "replace": replacement, "replace_all": replace_all, "replacements": len(matches)}

    def _word_live_insert_paragraphs(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        paragraphs = args.get("paragraphs")
        if not isinstance(paragraphs, list) or not all(isinstance(item, str) for item in paragraphs):
            raise ValidationError("word_live_insert_paragraphs.args.paragraphs must be an array of strings")
        position = args.get("position", "after")
        if args.get("target_text"):
            rng = _find_first_range(doc, args["target_text"])
        elif args.get("target_paragraph_index"):
            rng = doc.Paragraphs(int(args["target_paragraph_index"])).Range
        else:
            rng = _insert_range(doc, "end" if position in {"before", "after"} else position, args.get("bookmark"))
        text = "\r".join(_word_text(item) for item in paragraphs)
        if position == "before":
            rng.InsertBefore(text + "\r")
        else:
            if position == "after":
                _collapse_end(rng)
            rng.InsertAfter("\r" + text)
        if args.get("style"):
            for paragraph in _iter_collection(rng.Paragraphs):
                paragraph.Style = args["style"]
        return {"op": "word_live_insert_paragraphs", "paragraphs_inserted": len(paragraphs), "position": position}

    def _word_live_format_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        rng = _resolve_range(doc, args)
        font = rng.Font
        assignments = {
            "bold": "Bold", "italic": "Italic", "underline": "Underline",
            "strikethrough": "StrikeThrough", "font_name": "Name", "font_size": "Size",
        }
        for source, target in assignments.items():
            if source in args and args[source] is not None:
                value = args[source]
                if source == "underline":
                    value = 1 if value else 0
                setattr(font, target, value)
        if args.get("font_color") is not None:
            font.Color = _rgb(args["font_color"])
        if args.get("highlight_color") is not None:
            rng.HighlightColorIndex = args["highlight_color"]
        if args.get("style_name") is not None:
            rng.Style = args["style_name"]
        if args.get("paragraph_alignment") is not None:
            alignment = _enum(args["paragraph_alignment"], {"left": 0, "center": 1, "right": 2, "justify": 3}, "paragraph_alignment")
            for paragraph in _iter_collection(rng.Paragraphs):
                paragraph.Format.Alignment = alignment
        if args.get("page_break_before") is not None:
            for paragraph in _iter_collection(rng.Paragraphs):
                paragraph.Format.PageBreakBefore = bool(args["page_break_before"])
        return {"op": "word_live_format_text", "range": _range_info(rng)}

    def _word_live_apply_style(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        style_name = _required_string(args, "style_name")
        # Resolve the style before changing the range so a typo fails without
        # partially formatting the document.
        style = doc.Styles(style_name)
        rng = _resolve_range(doc, args)
        rng.Style = style
        return {"op": "word_live_apply_style", "style_name": style_name, "range": _range_info(rng)}

    def _word_live_add_hyperlink(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        url = _required_string(args, "url")
        rng = _resolve_range(doc, args)
        values = {"Anchor": rng, "Address": url}
        if args.get("sub_address"):
            values["SubAddress"] = args["sub_address"]
        if args.get("display_text") is not None:
            values["TextToDisplay"] = str(args["display_text"])
        link = doc.Hyperlinks.Add(**values)
        return {
            "op": "word_live_add_hyperlink",
            "address": str(_safe_value(link, "Address", url) or ""),
            "sub_address": str(_safe_value(link, "SubAddress", "") or ""),
            "text": str(_safe_value(_safe_value(link, "Range"), "Text", "")),
        }

    def _word_live_remove_hyperlink(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        removed = 0
        if args.get("hyperlink_index") is not None:
            index = int(args["hyperlink_index"])
            if index < 1 or index > _count(doc.Hyperlinks):
                raise ValidationError("hyperlink_index is out of range")
            doc.Hyperlinks(index).Delete()
            removed = 1
        else:
            rng = _resolve_range(doc, args)
            links = _safe_value(rng, "Hyperlinks", [])
            for link in list(_iter_collection(links)):
                link.Delete()
                removed += 1
        return {"op": "word_live_remove_hyperlink", "removed": removed}

    def _word_live_add_note(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        note_type = str(args.get("note_type", "footnote")).lower()
        if note_type not in {"footnote", "endnote"}:
            raise ValidationError("note_type must be 'footnote' or 'endnote'")
        rng = _resolve_range(doc, args)
        collection = doc.Footnotes if note_type == "footnote" else doc.Endnotes
        note = collection.Add(Range=rng, Text=_word_text(str(args["text"])))
        return {"op": "word_live_add_note", "note_type": note_type, "index": _count(collection), "text": str(args["text"])}

    def _word_live_update_fields(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        updated = 0
        for story in _story_ranges(doc):
            fields = _safe_value(story, "Fields")
            try:
                updated += _count(fields)
                fields.Update()
            except Exception:
                pass
        for table in _iter_collection(getattr(doc, "TablesOfContents", [])):
            try:
                table.Update()
            except Exception:
                pass
        return {"op": "word_live_update_fields", "fields": updated}

    def _word_live_insert_toc(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        rng = _insert_range(doc, args.get("position", "start"), args.get("bookmark"))
        kwargs = {
            "Range": rng,
            "UseHeadingStyles": bool(args.get("use_heading_styles", True)),
            "UpperHeadingLevel": int(args.get("upper_heading_level", 1)),
            "LowerHeadingLevel": int(args.get("lower_heading_level", 3)),
            "RightAlignPageNumbers": bool(args.get("right_align_page_numbers", True)),
            "IncludePageNumbers": bool(args.get("include_page_numbers", True)),
        }
        toc = doc.TablesOfContents.Add(**kwargs)
        return {"op": "word_live_insert_toc", "count": _count(doc.TablesOfContents), "range": _range_info(_safe_value(toc, "Range", rng))}

    def _word_live_set_content_control(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        title = _required_string(args, "title")
        value = str(args["value"])
        control = None
        created = False
        for item in _iter_collection(doc.ContentControls):
            if str(_safe_value(item, "Title", "")) == title or str(_safe_value(item, "Tag", "")) == title:
                control = item
                break
        if control is None:
            if not args.get("create_if_missing", True):
                raise ValidationError("content control was not found", details={"title": title})
            rng = _resolve_range(doc, args)
            control = doc.ContentControls.Add(1, rng)
            created = True
            control.Title = title
            if args.get("tag") is not None:
                control.Tag = str(args["tag"])
        control.Range.Text = value
        return {"op": "word_live_set_content_control", "title": title, "value": value, "created": created}

    def _word_live_remove_header_footer(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        position = str(args.get("position", "both")).lower()
        if position not in {"header", "footer", "both"}:
            raise ValidationError("position must be 'header', 'footer', or 'both'")
        section_index = args.get("section_index")
        sections = [_section(doc, args)] if section_index is not None else list(_iter_collection(doc.Sections))
        removed_shapes = 0
        cleared = []
        for section in sections:
            for area_name, collection in (("header", section.Headers), ("footer", section.Footers)):
                if position not in {area_name, "both"}:
                    continue
                for kind in (1, 2, 3):
                    area = collection(kind)
                    for shape in list(_iter_collection(_safe_value(area, "Shapes", []))):
                        shape.Delete()
                        removed_shapes += 1
                    rng = _safe_value(area, "Range")
                    if rng is not None:
                        try:
                            rng.Text = ""
                        except Exception:
                            try:
                                rng.Delete()
                            except Exception:
                                pass
                    cleared.append(area_name)
        return {
            "op": "word_live_remove_header_footer",
            "position": position,
            "sections": len(sections),
            "cleared": sorted(set(cleared)),
            "removed_shapes": removed_shapes,
        }

    def _word_live_export_pdf(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        output = _absolute_output(_required_string(args, "output_path"))
        if output.exists():
            raise ValidationError("PDF output already exists", details={"path": str(output)})
        output.parent.mkdir(parents=True, exist_ok=True)
        doc.ExportAsFixedFormat(
            OutputFileName=str(output),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            From=0,
            To=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
        if not output.is_file():
            raise ValidationError("Word did not create the requested PDF", details={"path": str(output)})
        return {"op": "word_live_export_pdf", "path": str(output), "bytes": output.stat().st_size}

    def _word_live_set_protection(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        protection_type = str(args.get("protection_type", "none")).lower()
        values = {"none": -1, "tracked_changes": 0, "comments": 1, "forms": 2, "read_only": 3}
        if protection_type not in values:
            raise ValidationError("protection_type must be none, tracked_changes, comments, forms, or read_only")
        password = args.get("password")
        if protection_type == "none":
            if _safe_value(doc, "ProtectionType", -1) != -1:
                doc.Unprotect(Password=str(password or ""))
        else:
            if _safe_value(doc, "ProtectionType", -1) != -1:
                doc.Unprotect(Password=str(password or ""))
            kwargs = {"Type": values[protection_type], "NoReset": True}
            if password is not None:
                kwargs["Password"] = str(password)
            doc.Protect(**kwargs)
        return {
            "op": "word_live_set_protection",
            "protection_type": protection_type,
            "protected": protection_type != "none",
        }

    def _word_live_compare_documents(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        source_path = Path(_required_string(args, "source_path"))
        output = _absolute_output(_required_string(args, "output_path"))
        if not source_path.is_file():
            raise ValidationError("compare source file does not exist", details={"source_path": str(source_path)})
        if output.exists():
            raise ValidationError("compare output already exists", details={"path": str(output)})
        source = None
        result = None
        try:
            source = app.Documents.Open(FileName=str(source_path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
            result = app.CompareDocuments(OriginalDocument=doc, RevisedDocument=source, Destination=2)
            if result is None:
                result = _find_document(app, str(source_path))
            output.parent.mkdir(parents=True, exist_ok=True)
            save_as2 = getattr(result, "SaveAs2", None)
            if callable(save_as2):
                save_as2(FileName=str(output), FileFormat=_file_format(output))
            else:
                result.SaveAs(FileName=str(output), FileFormat=_file_format(output))
            if not output.is_file():
                raise ValidationError("Word did not create the comparison output", details={"path": str(output)})
            return {"op": "word_live_compare_documents", "source_path": str(source_path), "path": str(output)}
        except ValidationError:
            raise
        except Exception as exc:
            raise ValidationError(
                "Word could not compare the documents",
                details={"source_path": str(source_path), "reason": str(exc)},
            ) from exc
        finally:
            if result is not None and result is not doc and result is not source:
                try:
                    result.Close(SaveChanges=False)
                except Exception:
                    pass
            if source is not None:
                try:
                    source.Close(SaveChanges=False)
                except Exception:
                    pass

    def _word_live_merge_document(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        source_path = Path(_required_string(args, "source_path"))
        if not source_path.is_file():
            raise ValidationError("merge source file does not exist", details={"source_path": str(source_path)})
        insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        insertion.InsertFile(FileName=str(source_path), ConfirmConversions=False, Link=False, Attachment=False)
        return {"op": "word_live_merge_document", "source_path": str(source_path), "merged": True}

    def _word_live_apply_list(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        paragraphs = _paragraph_range(doc, args)
        list_type = args.get("list_type", "bullet")
        for paragraph in _iter_collection(paragraphs.Paragraphs):
            listing = paragraph.Range.ListFormat
            if args.get("remove"):
                listing.RemoveNumbers()
            elif list_type == "bullet":
                listing.ApplyBulletDefault()
            elif list_type == "number":
                listing.ApplyNumberDefault()
            elif list_type == "multilevel":
                listing.ApplyOutlineNumberDefault()
            else:
                raise ValidationError("list_type must be bullet, number, or multilevel")
            if args.get("level") is not None:
                listing.ListLevelNumber = int(args["level"]) + 1
        return {"op": "word_live_apply_list", "list_type": list_type, "paragraphs": _count(paragraphs.Paragraphs)}

    def _word_live_setup_heading_numbering(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        template = doc.ListTemplates.Add(OutlineNumbered=True)
        styles = args.get("styles", {"Heading 1": 1, "Heading 2": 2, "Heading 3": 3})
        indexed_levels = {}
        for index in args.get("h1_paragraphs", []) or []:
            indexed_levels[int(index)] = 1
        for index in args.get("h2_paragraphs", []) or []:
            indexed_levels[int(index)] = 2
        for index, level in (args.get("heading_map", {}) or {}).items():
            if isinstance(index, str):
                for paragraph_index, paragraph in enumerate(_iter_collection(doc.Paragraphs), start=1):
                    if str(_safe_value(paragraph.Range, "Text", "")).strip().lower() == index.strip().lower():
                        indexed_levels[paragraph_index] = int(level)
            else:
                indexed_levels[int(index)] = int(level)
        applied = []
        for paragraph_index, paragraph in enumerate(_iter_collection(doc.Paragraphs), start=1):
            level = indexed_levels.get(paragraph_index)
            if level is None:
                for style, style_level in styles.items():
                    if _style_name(paragraph).lower() == str(style).lower():
                        level = int(style_level)
                        break
            if level is not None:
                paragraph.Range.ListFormat.ApplyListTemplateWithLevel(template, True, 0, 0, level)
                applied.append(paragraph_index)
                if args.get("font_name"):
                    paragraph.Range.Font.Name = args["font_name"]
                if args.get("bold") is not None:
                    paragraph.Range.Font.Bold = args["bold"]
                if args.get("alignment"):
                    paragraph.Format.Alignment = _enum(args["alignment"], {"left": 0, "center": 1, "right": 2, "justify": 3}, "alignment")
        return {"op": "word_live_setup_heading_numbering", "styles_applied": sorted(set(applied))}

    def _word_live_add_table(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        rows = int(args.get("rows", 0))
        columns = int(args.get("columns", args.get("cols", 0)))
        if rows < 1 or columns < 1:
            raise ValidationError("word_live_add_table requires positive rows and columns")
        rng = _insert_range(doc, args.get("position", "end"), args.get("bookmark"))
        table = doc.Tables.Add(Range=rng, NumRows=rows, NumColumns=columns)
        data = args.get("data", [])
        for row, values in enumerate(data, start=1):
            for column, value in enumerate(values, start=1):
                if row <= rows and column <= columns:
                    table.Cell(row, column).Range.Text = str(value)
        if args.get("style"):
            table.Style = args["style"]
        if args.get("autofit"):
            table.AutoFitBehavior(1 if args["autofit"] in {"window", "contents", True} else 0)
        return {"op": "word_live_add_table", "rows": rows, "columns": columns}

    def _word_live_format_table(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        table = _table(doc, args)
        table_alignment = args.get("alignment", args.get("table_alignment"))
        if table_alignment is not None:
            table.Rows.Alignment = _enum(table_alignment, {"left": 0, "center": 1, "right": 2}, "alignment")
        if args.get("autofit"):
            table.AutoFitBehavior(1)
        if args.get("style"):
            table.Style = args["style"]
        if args.get("borders") is not None:
            table.Borders.Enable = bool(args["borders"])
        if args.get("border_style") is not None:
            for border in _iter_collection(table.Borders):
                border.LineStyle = args["border_style"]
        for entry in args.get("cell_bold", []) or []:
            table.Cell(int(entry["row"]), int(entry["column"])).Range.Font.Bold = bool(entry.get("bold", True))
        for entry in args.get("cell_alignment", []) or []:
            table.Cell(int(entry["row"]), int(entry["column"])).Range.ParagraphFormat.Alignment = _enum(entry["alignment"], {"left": 0, "center": 1, "right": 2, "justify": 3}, "cell_alignment")
        for index, width in enumerate(args.get("column_widths", []) or [], start=1):
            table.Columns(index).Width = float(width)
        return {"op": "word_live_format_table", "table": int(args.get("table_index", 1))}

    def _word_live_modify_table(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        table = _table(doc, args)
        action = args.get("action", args.get("operation", "get_info"))
        row, column = int(args.get("row", 1)), int(args.get("column", args.get("col", 1)))
        if action == "get_info":
            return {"op": "word_live_modify_table", "action": action, "rows": _count(table.Rows), "columns": _count(table.Columns)}
        if action == "set_cell":
            table.Cell(row, column).Range.Text = str(args.get("text", ""))
        elif action == "set_row":
            values = args.get("cells", [])
            for index, text in enumerate(values, start=1):
                table.Cell(row, index).Range.Text = str(text)
        elif action == "set_range":
            for r_index, values in enumerate(args.get("rows", []), start=row):
                for c_index, text in enumerate(values, start=column):
                    table.Cell(r_index, c_index).Range.Text = str(text)
        elif action == "add_row":
            table.Rows.Add()
        elif action == "delete_row":
            table.Rows(row).Delete()
        elif action == "add_column":
            table.Columns.Add()
        elif action == "delete_column":
            table.Columns(column).Delete()
        elif action in {"merge", "merge_cells"}:
            table.Cell(row, column).Merge(table.Cell(int(args["end_row"]), int(args["end_column"])))
        elif action == "autofit":
            table.AutoFitBehavior(1)
        elif action == "delete_table":
            table.Delete()
        else:
            raise ValidationError("unsupported table action")
        return {"op": "word_live_modify_table", "action": action, "table": int(args.get("table_index", 1))}

    def _word_live_save(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        output = args.get("path", args.get("save_as"))
        if output:
            output_path = _absolute_output(output)
            if output_path.exists() and not args.get("overwrite", False):
                raise ValidationError("save path already exists; set overwrite explicitly", details={"path": str(output_path)})
            doc.SaveAs2(FileName=str(output_path), FileFormat=_file_format(output_path))
            return {"op": "word_live_save", "mode": "save-as", "path": str(output_path)}
        doc.Save()
        return {"op": "word_live_save", "mode": "in-place"}

    def _word_live_toggle_track_changes(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        previous = bool(doc.TrackRevisions)
        enabled = bool(args["enable"]) if "enable" in args else not previous
        doc.TrackRevisions = enabled
        return {"op": "word_live_toggle_track_changes", "previous": previous, "enabled": enabled}

    def _word_live_insert_image(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        image = Path(_required_string(args, "image_path"))
        if not image.is_file():
            raise ValidationError("image_path does not exist", details={"path": str(image)})
        position = args.get("position", "end")
        if args.get("paragraph_index"):
            rng = doc.Paragraphs(int(args["paragraph_index"])).Range
            if position == "after":
                _collapse_end(rng)
        else:
            rng = _insert_range(doc, position, args.get("bookmark"))
        shape = doc.InlineShapes.AddPicture(FileName=str(image), LinkToFile=False, SaveWithDocument=True, Range=rng)
        width = args.get("width_pt", args.get("width_inches", None))
        height = args.get("height_pt", args.get("height_inches", None))
        if width is not None:
            shape.Width = float(width) * (72 if "width_inches" in args and "width_pt" not in args else 1)
        if height is not None:
            shape.Height = float(height) * (72 if "height_inches" in args and "height_pt" not in args else 1)
        if args.get("alignment") is not None:
            rng.ParagraphFormat.Alignment = _enum(args["alignment"], {"left": 0, "center": 1, "right": 2}, "alignment")
        return {"op": "word_live_insert_image", "path": str(image), "width": _safe_value(shape, "Width"), "height": _safe_value(shape, "Height")}

    def _word_live_insert_cross_reference(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        reference_type = args.get("reference_type", args.get("ref_type", "heading"))
        item = args.get("reference_item", args.get("ref_item", 1))
        position = args.get("position", args.get("insert_position", "end"))
        if args.get("paragraph_index"):
            rng = doc.Paragraphs(int(args["paragraph_index"])).Range
            _collapse_end(rng)
        else:
            rng = _insert_range(doc, position, args.get("bookmark"))
        reference_kind = args.get("reference_kind", args.get("ref_kind", -1))
        if isinstance(reference_kind, str):
            reference_kind = {
                "content_text": -1,
                "page_number": 7,
                "position": 15,
                "number_full_context": -4,
                "number_relative_context": -2,
                "number_no_context": -3,
            }.get(reference_kind.strip().lower(), reference_kind)
        rng.InsertCrossReference(
            reference_type,
            int(reference_kind),
            item,
            bool(args.get("as_hyperlink", args.get("insert_as_hyperlink", True))),
            bool(args.get("include_position", False)),
            bool(args.get("separate_numbers", False)),
            str(args.get("separator_string", "")),
        )
        return {"op": "word_live_insert_cross_reference", "reference_type": reference_type, "reference_item": item}

    def _word_live_insert_equation(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        equation = _required_string(args, "equation")
        if args.get("paragraph_index"):
            rng = doc.Paragraphs(int(args["paragraph_index"])).Range
            _collapse_end(rng)
        else:
            rng = _insert_range(doc, args.get("position", "end"), args.get("bookmark"))
        rng.Text = equation
        doc.OMaths.Add(rng)
        # Dynamic pywin32 binds the return value of OMaths.Add as a generic
        # dispatch object named ``Add``. Re-fetching the collection item gives
        # us the actual OMath interface, including BuildUp.
        math = doc.OMaths.Item(doc.OMaths.Count)
        math.BuildUp()
        return {"op": "word_live_insert_equation", "equation": equation}

    def _word_live_take_snapshot(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        self._snapshots[_doc_key(doc)] = self._snapshot_paragraphs(doc)
        return {"op": "word_live_take_snapshot", "paragraphs": len(self._snapshots[_doc_key(doc)])}

    def _word_live_get_diff(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        before = self._snapshots.get(_doc_key(doc))
        if before is None:
            raise ValidationError("no snapshot exists for this open document")
        after = self._snapshot_paragraphs(doc)
        changes = []
        for index in range(max(len(before), len(after))):
            old = before[index]["text"] if index < len(before) else None
            new = after[index]["text"] if index < len(after) else None
            if old != new:
                changes.append({"index": index + 1, "before": old, "after": new})
        return {"op": "word_live_get_diff", "changes": changes}

    def _word_live_snapshot_status(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        exists = _doc_key(doc) in self._snapshots
        return {"op": "word_live_snapshot_status", "exists": exists}

    def _word_live_get_page_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        page = int(args.get("page", 1))
        end_page = int(args.get("end_page", page))
        total_pages = max(1, int(doc.ComputeStatistics(2)))
        if page < 1 or end_page < page:
            raise ValidationError("page must be positive and end_page must not precede page")
        if page > total_pages:
            text = ""
        else:
            start = _page_boundary(doc, page, total_pages)
            end = _page_boundary(doc, min(end_page + 1, total_pages + 1), total_pages)
            text = str(_safe_value(doc.Range(start, end), "Text", ""))
        return {"op": "word_live_get_page_text", "page": page, "end_page": end_page, "text": text}

    def _word_live_get_paragraph_format(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        paragraphs = _paragraph_range(doc, args).Paragraphs
        values = []
        for index, paragraph in enumerate(_iter_collection(paragraphs), start=1):
            font = paragraph.Range.Font
            fmt = paragraph.Format
            values.append({"index": index, "text": str(_safe_value(paragraph.Range, "Text", "")).rstrip("\r\a"), "style": _style_name(paragraph), "font": {"name": _safe_value(font, "Name"), "size": _safe_value(font, "Size"), "bold": _safe_value(font, "Bold"), "italic": _safe_value(font, "Italic")}, "paragraph": {"alignment": _safe_value(fmt, "Alignment"), "space_before": _safe_value(fmt, "SpaceBefore"), "space_after": _safe_value(fmt, "SpaceAfter"), "keep_with_next": _safe_value(fmt, "KeepWithNext"), "keep_together": _safe_value(fmt, "KeepTogether")}})
        return {"op": "word_live_get_paragraph_format", "paragraphs": values}

    def _word_live_find_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        text = _required_string(args, "find_text") if "find_text" in args else _required_string(args, "search_text")
        rng = doc.Content.Duplicate
        finder = _configure_find(rng.Find, text, args)
        max_results = int(args.get("max_results", 50))
        matches = []
        while len(matches) < max_results and finder.Execute():
            start, end = int(rng.Start), int(rng.End)
            context = str(_safe_value(doc.Range(max(0, start - int(args.get("context_chars", 60))), min(doc.Content.End, end + int(args.get("context_chars", 60)))), "Text", ""))
            matches.append({"text": str(rng.Text), "start": start, "end": end, "context": context})
            rng.Collapse(0)
            finder = rng.Find
            finder = _configure_find(finder, text, args)
        return {"op": "word_live_find_text", "find": text, "matches": matches}

    def _word_live_get_undo_history(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        return {"op": "word_live_get_undo_history", "entries": list(self._history)}

    def _word_live_list_cross_reference_items(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        bookmarks = [{"name": str(_safe_value(item, "Name", "")), "range": _range_info(item.Range)} for item in _iter_collection(doc.Bookmarks)]
        headings = []
        for paragraph in _iter_collection(doc.Paragraphs):
            if _style_name(paragraph).lower().startswith("heading"):
                headings.append({"text": str(_safe_value(paragraph.Range, "Text", "")).rstrip("\r\a"), "style": _style_name(paragraph)})
        return {"op": "word_live_list_cross_reference_items", "bookmarks": bookmarks, "headings": headings}

    def _word_live_diagnose_layout(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        warnings = []
        paragraphs = list(_iter_collection(doc.Paragraphs))
        for index, paragraph in enumerate(paragraphs, start=1):
            fmt = paragraph.Format
            if _safe_value(fmt, "KeepWithNext", False) and index == len(paragraphs):
                warnings.append({"paragraph": index, "issue": "keep_with_next_on_final_paragraph"})
            if _safe_value(fmt, "KeepTogether", False) and len(str(_safe_value(paragraph.Range, "Text", ""))) > 1000:
                warnings.append({"paragraph": index, "issue": "large_keep_together_paragraph"})
        return {"op": "word_live_diagnose_layout", "warnings": warnings}

    def _word_live_get_comments(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        comments = []
        for index, comment in enumerate(_iter_collection(doc.Comments), start=1):
            comments.append(_comment_dict(comment, index))
        return {"op": "word_live_get_comments", "comments": comments}

    def _word_live_add_comment(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        rng = _resolve_range(doc, args)
        # Word's Comments.Add COM signature is only (Range, Text); Author and
        # Initials are document/application metadata, not accepted parameters
        # by the current Word type library.
        comment = doc.Comments.Add(rng, args.get("comment_text", args.get("text", "")))
        return {"op": "word_live_add_comment", "comment": _comment_dict(comment, _count(doc.Comments))}

    def _word_live_list_revisions(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        revisions = []
        for index, revision in enumerate(_iter_collection(doc.Revisions), start=1):
            revisions.append({"id": index, "type": _safe_value(revision, "Type"), "author": _safe_value(revision, "Author"), "date": str(_safe_value(revision, "Date", "")), "text": str(_safe_value(revision.Range, "Text", ""))})
        return {"op": "word_live_list_revisions", "revision_count": len(revisions), "revisions": revisions}

    def _word_live_reply_to_comment(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        comment = _comment(doc, args)
        reply = comment.Replies.Add(comment.Range, args.get("reply_text", args.get("text", "")))
        return {"op": "word_live_reply_to_comment", "comment": _comment_dict(reply, int(args.get("comment_index", 1)))}

    def _word_live_resolve_comment(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        comment = _comment(doc, args)
        resolved = bool(args.get("resolved", args.get("resolve", True)))
        try:
            comment.Done = resolved
        except Exception:
            # Modern Comments expose Done for compatibility, but current Word
            # builds reject setting it through COM. Persist the same state in
            # the OOXML threaded-comment extension after the write commit.
            self._pending_comment_package_edits.setdefault(_doc_key(doc), []).append(
                {"comment_index": int(args.get("comment_index", 1)), "resolved": resolved}
            )
        return {"op": "word_live_resolve_comment", "resolved": resolved}

    def _commit_pending_comment_edits(self, app: Any, doc: Any, path: Path, *, reopen: bool) -> None:
        edits = self._pending_comment_package_edits.pop(_doc_key(doc), [])
        if not edits:
            return
        if not path.is_file():
            raise ValidationError("cannot update comment resolution: Word file does not exist", details={"path": str(path)})
        if reopen:
            doc.Close(SaveChanges=False)
        try:
            _set_comment_resolution_in_package(path, edits)
        finally:
            if reopen:
                app.Documents.Open(FileName=str(path), ReadOnly=False, AddToRecentFiles=False)

    def _word_live_delete_comment(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        comment = _comment(doc, args)
        try:
            comment.Delete()
        except Exception:
            delete_thread = getattr(comment, "DeleteRecursively", None)
            if not callable(delete_thread):
                raise
            delete_thread()
        return {"op": "word_live_delete_comment", "deleted": True}

    def _word_live_accept_revisions(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        return self._revision_action(doc, args, "Accept")

    def _word_live_reject_revisions(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        return self._revision_action(doc, args, "Reject")

    def _revision_action(self, doc: Any, args: dict[str, Any], action: str) -> dict[str, Any]:
        revisions = doc.Revisions
        ids = args.get("revision_ids")
        author = args.get("author")
        count = 0
        if isinstance(ids, list):
            for revision_id in sorted((int(item) for item in ids), reverse=True):
                if 1 <= revision_id <= _count(revisions):
                    getattr(revisions(revision_id), action)()
                    count += 1
        elif author:
            for index in range(_count(revisions), 0, -1):
                revision = revisions(index)
                if str(_safe_value(revision, "Author", "")) == author:
                    getattr(revision, action)()
                    count += 1
        else:
            count = _count(revisions)
            getattr(doc, f"{action}AllRevisions")()
        return {"op": f"word_live_{action.lower()}_revisions", "count": count}

    def _word_live_set_page_layout(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        section = _section(doc, args)
        page = section.PageSetup
        if args.get("orientation") is not None:
            page.Orientation = _enum(args["orientation"], {"portrait": 0, "landscape": 1}, "orientation")
        for source, target in (("page_width_inches", "PageWidth"), ("page_height_inches", "PageHeight"), ("top_margin_inches", "TopMargin"), ("bottom_margin_inches", "BottomMargin"), ("left_margin_inches", "LeftMargin"), ("right_margin_inches", "RightMargin")):
            if args.get(source) is not None:
                setattr(page, target, float(args[source]) * 72)
        return {"op": "word_live_set_page_layout", "section": int(args.get("section_index", 1))}

    def _word_live_add_header_footer(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        section = _section(doc, args)
        alignments = {"left": 0, "center": 1, "right": 2}
        added = []
        if args.get("header_text") is not None:
            header = section.Headers(1)
            header.Range.Text = args["header_text"]
            header.Range.ParagraphFormat.Alignment = _enum(args.get("header_alignment", "center"), alignments, "header_alignment")
            added.append("header")
        if args.get("footer_text") is not None:
            footer = section.Footers(1)
            footer.Range.Text = args["footer_text"]
            footer.Range.ParagraphFormat.Alignment = _enum(args.get("footer_alignment", "center"), alignments, "footer_alignment")
            added.append("footer")
        return {"op": "word_live_add_header_footer", "section": int(args.get("section_index", 1)), "added": added}

    def _word_live_add_page_numbers(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        section = _section(doc, args)
        position = args.get("position", "footer")
        target = section.Headers(1) if position == "header" else section.Footers(1)
        existing = str(_safe_value(target.Range, "Text", "")).replace("\r", "").replace("\a", "").strip()
        if existing:
            # Keep authored header/footer text in its own paragraph. Word's
            # PageNumbers collection positions fields independently and can
            # overlap centered authored text, so use explicit PAGE/NUMPAGES
            # fields in a dedicated paragraph instead.
            target.Range.InsertAfter("\r")
        paragraph = target.Range.Paragraphs(target.Range.Paragraphs.Count).Range.Duplicate
        paragraph.Collapse(1)
        if args.get("prefix"):
            paragraph.InsertAfter(str(args["prefix"]))
            paragraph.Collapse(0)
        doc.Fields.Add(Range=paragraph, Type=33)  # wdFieldPage
        if args.get("include_total"):
            paragraph = target.Range.Paragraphs(target.Range.Paragraphs.Count).Range.Duplicate
            paragraph.Collapse(0)
            paragraph.InsertAfter(" / ")
            paragraph.Collapse(0)
            doc.Fields.Add(Range=paragraph, Type=26)  # wdFieldNumPages
        if args.get("suffix"):
            paragraph = target.Range.Paragraphs(target.Range.Paragraphs.Count).Range.Duplicate
            paragraph.Collapse(0)
            paragraph.InsertAfter(str(args["suffix"]))
        target.Range.Paragraphs(target.Range.Paragraphs.Count).Range.ParagraphFormat.Alignment = _enum(
            args.get("alignment", "center"), {"left": 0, "center": 1, "right": 2}, "alignment"
        )
        doc.Fields.Update()
        return {
            "op": "word_live_add_page_numbers",
            "section": int(args.get("section_index", 1)),
            "position": position,
            "fields": int(target.Range.Fields.Count),
        }

    def _word_live_add_section_break(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        kind = _enum(args.get("break_type", "new_page"), {"new_page": 2, "continuous": 3, "even_page": 4, "odd_page": 5}, "break_type")
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertBreak(Type=kind)
        return {"op": "word_live_add_section_break", "break_type": args.get("break_type", "new_page"), "sections": _count(doc.Sections)}

    def _word_live_set_paragraph_spacing(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        rule = {"single": 0, "1.5_lines": 1, "double": 2, "at_least": 3, "exactly": 4, "multiple": 5}
        paragraphs = _paragraph_range(doc, args).Paragraphs
        for paragraph in _iter_collection(paragraphs):
            fmt = paragraph.Format
            for source, target in (("space_before_pt", "SpaceBefore"), ("space_after_pt", "SpaceAfter"), ("line_spacing", "LineSpacing"), ("keep_with_next", "KeepWithNext"), ("keep_together", "KeepTogether")):
                if args.get(source) is not None:
                    setattr(fmt, target, args[source])
            if args.get("line_spacing_rule") is not None:
                fmt.LineSpacingRule = _enum(args["line_spacing_rule"], rule, "line_spacing_rule")
            if args.get("alignment") is not None:
                fmt.Alignment = _enum(args["alignment"], {"left": 0, "center": 1, "right": 2, "justify": 3}, "alignment")
        return {"op": "word_live_set_paragraph_spacing", "paragraphs_affected": _count(paragraphs)}

    def _word_live_add_bookmark(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        name = _required_string(args, "bookmark_name")
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", name):
            raise ValidationError("bookmark_name must contain letters, numbers, and underscores and start with a letter")
        paragraph = int(args.get("paragraph_index", 1))
        if paragraph < 1 or paragraph > _count(doc.Paragraphs):
            raise ValidationError("paragraph_index is out of range")
        doc.Bookmarks.Add(name, doc.Paragraphs(paragraph).Range)
        return {"op": "word_live_add_bookmark", "bookmark_name": name, "paragraph_index": paragraph}

    def _word_live_add_watermark(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        section = _section(doc, args)
        shape = section.Headers(1).Shapes.AddTextEffect(0, args.get("text", "DRAFT"), "Calibri", args.get("font_size", 72), False, False, 0, 0)
        shape.Fill.ForeColor.RGB = _rgb(args.get("font_color", "C0C0C0"))
        shape.Fill.Transparency = 0.5
        shape.Line.Visible = False
        shape.Rotation = args.get("rotation", -45)
        shape.WrapFormat.Type = 3
        shape.Left = -999995
        shape.Top = -999995
        return {"op": "word_live_add_watermark", "text": args.get("text", "DRAFT"), "section": int(args.get("section_index", 1))}

    def _word_live_remove_watermark(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        """Remove exact-text WordArt watermarks from header/footer stories.

        Word rejects deleting a header WordArt object while revision tracking is
        enabled.  Temporarily disabling tracking for this narrowly scoped
        object deletion keeps the operation from creating a spurious tracked
        change; the user's original tracking setting is restored in ``finally``.
        """

        text = args.get("text", "DRAFT")
        if not isinstance(text, str) or not text.strip():
            raise ValidationError("text must be a non-empty string")
        position = str(args.get("position", "both")).lower()
        if position not in {"header", "footer", "both"}:
            raise ValidationError("position must be 'header', 'footer', or 'both'")

        section_index = args.get("section_index")
        sections = [_section(doc, args)] if section_index is not None else list(_iter_collection(doc.Sections))
        previous_tracking = _safe_value(doc, "TrackRevisions", None)
        view = None
        removed = 0
        inspected = 0
        try:
            try:
                doc.Activate()
            except Exception:
                pass
            if bool(previous_tracking):
                doc.TrackRevisions = False
            for section in sections:
                for kind in (1, 2, 3):
                    areas = []
                    if position in {"header", "both"}:
                        areas.append(("header", section.Headers(kind)))
                    if position in {"footer", "both"}:
                        areas.append(("footer", section.Footers(kind)))
                    for area_name, area in areas:
                        view = _enter_header_footer_view(doc, area_name)
                        while True:
                            matches = []
                            for shape in list(_iter_collection(_safe_value(area, "Shapes", []))):
                                inspected += 1
                                if _shape_watermark_text(shape).casefold() == text.strip().casefold():
                                    matches.append(shape)
                            if not matches:
                                break
                            for shape in matches:
                                try:
                                    shape.Delete()
                                except Exception as exc:
                                    raise ValidationError(
                                        "Word could not remove the matching watermark",
                                        details={"text": text, "position": area_name, "section": section_index or "all"},
                                    ) from exc
                                removed += 1
        finally:
            if previous_tracking is not None:
                try:
                    doc.TrackRevisions = previous_tracking
                except Exception:
                    pass
            if view is not None:
                try:
                    view.SeekView = 0
                except Exception:
                    pass
        return {
            "op": "word_live_remove_watermark",
            "text": text,
            "position": position,
            "section": section_index or "all",
            "inspected": inspected,
            "removed": removed,
            "track_changes_restored": previous_tracking is not None,
        }

    def _word_live_undo(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        count = max(1, int(args.get("count", args.get("times", 1))))
        for _ in range(count):
            doc.Undo()
        return {"op": "word_live_undo", "count": count}

    def _word_screen_capture(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        output = _absolute_output(_required_string(args, "output_path"))
        try:
            app.Visible = True
        except Exception:
            pass
        try:
            doc.Activate()
        except Exception:
            pass
        window = doc.ActiveWindow
        try:
            window.Activate()
        except Exception:
            pass
        try:
            app.ScreenRefresh()
        except Exception:
            pass
        time.sleep(0.15)
        hwnd = int(window.Hwnd)
        self._screenshotter(hwnd, output)
        return {"op": "word_screen_capture", "output_path": str(output)}

    def _word_live_set_core_properties(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        props = doc.BuiltInDocumentProperties
        changed = []
        for source, name in (("title", "Title"), ("subject", "Subject"), ("author", "Author"), ("keywords", "Keywords"), ("comments", "Comments"), ("category", "Category"), ("company", "Company"), ("manager", "Manager"), ("last_author", "Last Author") ):
            if source in args and args[source] is not None:
                props(name).Value = args[source]
                changed.append(source)
        return {"op": "word_live_set_core_properties", "changed": changed}

    def _snapshot_paragraphs(self, doc: Any) -> list[dict[str, Any]]:
        return [{"index": index, "text": str(_safe_value(paragraph.Range, "Text", "")).rstrip("\r\a"), "style": _style_name(paragraph)} for index, paragraph in enumerate(_iter_collection(doc.Paragraphs), start=1)]


def _operation_args(operation: dict[str, Any]) -> dict[str, Any]:
    if operation["op"] == "replace_text":
        return {"find_text": operation["find"], "replace_text": operation["replace"], "replace_all": operation.get("occurrence", "all") == "all"}
    return dict(operation.get("args", {}))


def _default_word_app():
    try:
        from win32com.client import Dispatch, GetActiveObject
    except ImportError as exc:
        raise ImportError("pywin32 is required for --backend com") from exc
    try:
        return GetActiveObject("Word.Application")
    except Exception:
        return Dispatch("Word.Application")


def _find_document(app: Any, target: str) -> Any:
    wanted = Path(target).resolve()
    for doc in _iter_collection(getattr(app, "Documents", [])):
        candidates = {str(_safe_value(doc, "FullName", "")), str(_safe_value(doc, "Name", ""))}
        if str(wanted) in candidates or target in candidates or Path(str(_safe_value(doc, "FullName", ""))).resolve() == wanted:
            return doc
    raise ValidationError("target Word document is not open", details={"target": target})


def _find_or_open_for_inspect(app: Any, target: Path) -> tuple[Any, bool]:
    try:
        return _find_document(app, str(target)), False
    except ValidationError:
        if not target.is_file():
            raise
        try:
            document = app.Documents.Open(
                FileName=str(target),
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
        except Exception as exc:
            raise ValidationError(
                "Word could not open the target for inspection",
                details={"target": str(target), "reason": str(exc)},
            ) from exc
        return document, True


def _render_document(document: Any, render_dir: Path) -> dict[str, Any]:
    output_dir = render_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    existing = sorted(str(path) for path in output_dir.glob("page-*.png"))
    if existing:
        raise ValidationError("render output already exists", details={"paths": existing})
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise BackendUnavailableError("word-render")

    with tempfile.TemporaryDirectory(prefix="white-collar-word-render-") as temporary:
        temporary_dir = Path(temporary)
        pdf_path = temporary_dir / "document.pdf"
        try:
            document.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                From=0,
                To=0,
                Item=0,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=0,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        except Exception as exc:
            raise ValidationError(
                "Word could not export the document for rendering",
                details={"target": str(_safe_value(document, "FullName", "")), "reason": str(exc)},
            ) from exc
        if not pdf_path.is_file():
            raise ValidationError("Word did not create the temporary PDF for rendering", details={"target": str(_safe_value(document, "FullName", ""))})

        prefix = temporary_dir / "page"
        try:
            completed = subprocess.run(
                [pdftoppm, "-png", "-r", "144", str(pdf_path), str(prefix)],
                check=False,
                capture_output=True,
                text=True,
            )
        except OSError as exc:
            raise BackendUnavailableError("word-render") from exc
        if completed.returncode != 0:
            raise ValidationError(
                "the PDF renderer could not rasterize the Word document",
                details={"returncode": completed.returncode, "stderr": completed.stderr.strip()},
            )
        rendered = sorted(
            temporary_dir.glob("page-*.png"),
            key=lambda path: int(path.stem.rsplit("-", 1)[1]),
        )
        if not rendered:
            raise ValidationError("the PDF renderer produced no Word pages")

        files: list[str] = []
        for index, source in enumerate(rendered, start=1):
            destination = output_dir / f"page-{index}.png"
            shutil.copy2(source, destination)
            files.append(str(destination))
    return {"directory": str(output_dir), "format": "png", "pages": len(files), "files": files}


def _undo_record(app: Any, name: str):
    undo = getattr(app, "UndoRecord", None)
    if undo is None:
        return contextlib.nullcontext()
    @contextlib.contextmanager
    def record():
        undo.StartCustomRecord(name)
        try:
            yield
        finally:
            undo.EndCustomRecord()
    return record()


@contextlib.contextmanager
def _tracking(doc: Any, enabled: bool | None) -> Iterator[None]:
    if enabled is None:
        yield
        return
    previous = bool(doc.TrackRevisions)
    doc.TrackRevisions = True
    try:
        yield
    finally:
        doc.TrackRevisions = previous


def _resolve_range(doc: Any, args: dict[str, Any]) -> Any:
    if "start" in args and "end" in args:
        return doc.Range(int(args["start"]), int(args["end"]))
    if "start_paragraph" in args:
        start = int(args["start_paragraph"])
        end = int(args.get("end_paragraph", start))
        return doc.Range(doc.Paragraphs(start).Range.Start, doc.Paragraphs(end).Range.End)
    if "paragraph_index" in args:
        return doc.Paragraphs(int(args["paragraph_index"])).Range
    if "bookmark" in args:
        if not doc.Bookmarks.Exists(args["bookmark"]):
            raise ValidationError("bookmark not found", details={"bookmark": args["bookmark"]})
        return doc.Bookmarks(args["bookmark"]).Range
    if "target_text" in args:
        return _find_first_range(doc, _required_string(args, "target_text"))
    return doc.Content.Duplicate


def _find_first_range(doc: Any, text: str) -> Any:
    rng = doc.Content.Duplicate
    finder = _configure_find(rng.Find, text, {})
    if not finder.Execute():
        raise ValidationError("target text was not found", details={"target_text": text})
    return rng


def _insert_range(doc: Any, position: Any, bookmark: str | None) -> Any:
    if bookmark:
        if not doc.Bookmarks.Exists(bookmark):
            raise ValidationError("bookmark not found", details={"bookmark": bookmark})
        return doc.Bookmarks(bookmark).Range
    if position == "start":
        return doc.Range(0, 0)
    if position == "end":
        end = doc.Content.End - 1
        return doc.Range(end, end)
    if position == "cursor":
        return doc.Application.Selection.Range
    try:
        offset = int(position)
    except (TypeError, ValueError) as exc:
        raise ValidationError("position must be start, end, cursor, or a character offset") from exc
    return doc.Range(offset, offset)


def _paragraph_range(doc: Any, args: dict[str, Any]) -> Any:
    start = int(args.get("start_paragraph", args.get("paragraph_index", 1)))
    end = int(args.get("end_paragraph", start))
    if start < 1 or end < start or end > _count(doc.Paragraphs):
        raise ValidationError("paragraph range is out of bounds")
    return doc.Range(doc.Paragraphs(start).Range.Start, doc.Paragraphs(end).Range.End)


def _table(doc: Any, args: dict[str, Any]) -> Any:
    index = int(args.get("table_index", 1))
    if index == -1:
        index = _count(doc.Tables)
    if index < 1 or index > _count(doc.Tables):
        raise ValidationError("table_index is out of range")
    return doc.Tables(index)


def _section(doc: Any, args: dict[str, Any]) -> Any:
    index = int(args.get("section_index", 1))
    if index < 1 or index > _count(doc.Sections):
        raise ValidationError("section_index is out of range")
    return doc.Sections(index)


def _comment(doc: Any, args: dict[str, Any]) -> Any:
    index = int(args.get("comment_index", 1))
    if index < 1 or index > _count(doc.Comments):
        raise ValidationError("comment_index is out of range")
    return doc.Comments(index)


def _comment_dict(comment: Any, index: int) -> dict[str, Any]:
    comment_range = _safe_value(comment, "Range")
    return {"index": index, "author": _safe_value(comment, "Author", ""), "text": str(_safe_value(comment_range, "Text", "")), "date": str(_safe_value(comment, "Date", ""))}


def _iter_collection(collection: Any):
    if collection is None:
        return iter(())
    try:
        return iter(collection)
    except TypeError:
        count = _count(collection)
        return (collection(index) for index in range(1, count + 1))


def _story_ranges(doc: Any):
    """Yield each Word story, including linked continuation ranges."""

    stories = _safe_value(doc, "StoryRanges", [])
    for story in _iter_collection(stories):
        current = story
        while current is not None:
            yield current
            current = _safe_value(current, "NextStoryRange")


def _count(collection: Any) -> int:
    try:
        return int(collection.Count)
    except (AttributeError, TypeError, ValueError):
        try:
            return len(collection)
        except TypeError:
            return 0


def _safe_value(obj: Any, attr: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, attr)
        return value() if callable(value) and attr == "__call__" else value
    except Exception:
        return default


def _shape_watermark_text(shape: Any) -> str:
    text_effect = _safe_value(shape, "TextEffect")
    value = _safe_value(text_effect, "Text") if text_effect is not None else None
    if value is None or not str(value).strip():
        text_frame = _safe_value(shape, "TextFrame")
        text_range = _safe_value(text_frame, "TextRange") if text_frame is not None else None
        value = _safe_value(text_range, "Text", "") if text_range is not None else ""
    return str(value or "").replace("\r", "").replace("\a", "").strip()


def _enter_header_footer_view(doc: Any, position: str) -> Any:
    """Put Word in a story view when COM requires it for shape deletion."""

    try:
        window = _safe_value(doc, "ActiveWindow")
        if window is None:
            windows = _safe_value(doc, "Windows")
            if _count(windows):
                window = windows(1)
        view = _safe_value(window, "View") if window is not None else None
        if view is not None:
            # Word accepts the current header/footer seek views for all linked
            # header/footer stories exposed through the section collections.
            view.SeekView = 9 if position == "header" else 10
            return view
    except Exception:
        pass
    return None


def _style_name(paragraph: Any) -> str | None:
    style = _safe_value(paragraph, "Style")
    return str(_safe_value(style, "NameLocal", style)) if style is not None else None


def _range_info(rng: Any) -> dict[str, Any]:
    return {"start": _safe_value(rng, "Start"), "end": _safe_value(rng, "End")}


def _page_boundary(doc: Any, page: int, total_pages: int) -> int:
    content_start = int(doc.Content.Start)
    content_end = int(doc.Content.End)
    if page <= 1:
        return content_start
    if page > total_pages:
        return content_end

    low, high = content_start, max(content_start, content_end - 1)
    while low < high:
        middle = (low + high) // 2
        probe = doc.Range(middle, min(middle + 1, content_end))
        if int(probe.Information(3)) >= page:
            high = middle
        else:
            low = middle + 1
    return low


def _collapse_end(rng: Any) -> None:
    try:
        rng.Collapse(0)
    except Exception:
        pass


def _required_string(args: dict[str, Any], name: str) -> str:
    value = args.get(name)
    if not isinstance(value, str) or not value:
        raise ValidationError(f"{name} must be a non-empty string")
    return value


def _configure_find(finder: Any, text: str, args: dict[str, Any]) -> Any:
    finder.ClearFormatting()
    finder.Text = text
    finder.Forward = True
    finder.Wrap = 0
    finder.Format = False
    for source, target in (("match_case", "MatchCase"), ("match_whole_word", "MatchWholeWord"), ("whole_word", "MatchWholeWord"), ("use_wildcards", "MatchWildcards")):
        if source in args:
            setattr(finder, target, bool(args[source]))
    return finder


def _word_text(value: str) -> str:
    return value.replace("\\r\\n", "\r").replace("\\r", "\r").replace("\\n", "\r")


def _enum(value: Any, values: dict[str, int], name: str) -> int:
    key = str(value).lower()
    if key not in values:
        raise ValidationError(f"invalid {name}: {value}", details={"allowed": sorted(values)})
    return values[key]


def _rgb(value: str) -> int:
    color = str(value).lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", color):
        raise ValidationError("color must be a six-digit RGB hex value")
    return int(color[0:2], 16) + (int(color[2:4], 16) << 8) + (int(color[4:6], 16) << 16)


def _absolute_output(value: str) -> Path:
    path = Path(value)
    if not path.is_absolute():
        raise ValidationError("output path must be absolute")
    return path


def _file_format(path: Path) -> int:
    return {".docx": 12, ".pdf": 17, ".rtf": 6, ".txt": 2}.get(path.suffix.lower(), 12)


def _set_comment_resolution_in_package(path: Path, edits: list[dict[str, Any]]) -> None:
    comments_extended = "word/commentsExtended.xml"
    comment_pattern = re.compile(rb"<w15:commentEx\b[^>]*?\/>")
    temporary_name = None
    try:
        with zipfile.ZipFile(path, "r") as source_zip:
            if comments_extended not in source_zip.namelist():
                raise ValidationError(
                    "Word file has no threaded-comment extension; cannot persist modern comment resolution",
                    details={"path": str(path)},
                )
            infos = source_zip.infolist()
            payloads = {info.filename: source_zip.read(info.filename) for info in infos}
            payload = payloads[comments_extended]
            entries = list(comment_pattern.finditer(payload))
            for edit in edits:
                index = int(edit["comment_index"])
                if index < 1 or index > len(entries):
                    raise ValidationError(
                        "comment_index is out of range in threaded-comment extension",
                        details={"comment_index": index, "count": len(entries)},
                    )
                match = entries[index - 1]
                element = match.group(0)
                value = b"1" if edit["resolved"] else b"0"
                if re.search(rb"w15:done=\"[^\"]*\"", element):
                    element = re.sub(rb"w15:done=\"[^\"]*\"", b'w15:done="' + value + b'"', element, count=1)
                else:
                    element = element[:-2] + b' w15:done="' + value + b'"/>'
                payload = payload[: match.start()] + element + payload[match.end() :]
                entries = list(comment_pattern.finditer(payload))
            payloads[comments_extended] = payload

        with tempfile.NamedTemporaryFile(
            mode="wb", suffix=path.suffix, prefix=f".{path.stem}-", dir=path.parent, delete=False
        ) as temporary:
            temporary_name = Path(temporary.name)
        with zipfile.ZipFile(temporary_name, "w") as destination_zip:
            for info in infos:
                destination_zip.writestr(info, payloads[info.filename])
        os.replace(temporary_name, path)
        temporary_name = None
    finally:
        if temporary_name is not None:
            temporary_name.unlink(missing_ok=True)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _doc_key(doc: Any) -> str:
    return str(_safe_value(doc, "FullName", _safe_value(doc, "Name", id(doc))))
