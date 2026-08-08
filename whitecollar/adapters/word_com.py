from __future__ import annotations

import contextlib
import datetime as dt
import hashlib
import os
import re
import time
from pathlib import Path
from typing import Any, Callable, Iterator

from ..errors import BackendUnavailableError, ValidationError
from ..models import Plan
from ..word_ops import WORD_COM_MUTATING_OPERATIONS, WORD_COM_OPERATIONS


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
        self._screenshotter = screenshotter or _default_screenshot
        self._snapshots: dict[str, list[dict[str, Any]]] = {}
        self._history: list[dict[str, Any]] = []

    def inspect(self, target: Path) -> dict[str, Any]:
        app = self._get_app()
        doc = _find_document(app, str(target))
        return self._get_info(doc) | {"backend": "word-com", "text": self._get_text(doc)["paragraphs"]}

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
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
        if not dry_run and doc is not None and plan.write.mode != "none":
            self._commit_write(doc, plan)
        return {"backend": "word-com", "written": not dry_run, "operations": operations}

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
        if not callable(save_copy_as):
            raise ValidationError("Word COM document does not support SaveCopyAs; cannot create required snapshot")
        save_copy_as(str(snapshot))

    def _commit_write(self, doc: Any, plan: Plan) -> None:
        if plan.write.mode == "in-place":
            doc.Save()
            return
        output = Path(plan.write.path or "")
        if output.exists():
            raise ValidationError("save-as path already exists", details={"path": str(output)})
        output.parent.mkdir(parents=True, exist_ok=True)
        save_copy_as = getattr(doc, "SaveCopyAs", None)
        if callable(save_copy_as):
            save_copy_as(str(output))
        else:
            doc.SaveAs2(FileName=str(output), FileFormat=_file_format(output))

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

    def _get_text(self, doc: Any) -> dict[str, Any]:
        return self._word_live_get_text(None, doc, {})

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
        with _tracking(doc, bool(args.get("track_changes"))):
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
        with _tracking(doc, bool(args.get("track_changes"))):
            deleted = str(_safe_value(rng, "Text", ""))
            rng.Delete()
        return {"op": "word_live_delete_text", "characters": len(deleted)}

    def _word_live_replace_text(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        find_text = _required_string(args, "find_text")
        replacement = str(args.get("replace_text", ""))
        finder = doc.Content.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        finder.Text = find_text
        finder.Replacement.Text = replacement
        finder.Forward = True
        finder.Wrap = 1
        finder.Format = False
        for prop in ("MatchCase", "MatchWholeWord", "MatchWildcards"):
            if prop in args:
                setattr(finder, prop, bool(args[prop]))
        replace_all = bool(args.get("replace_all", True))
        with _tracking(doc, bool(args.get("track_changes"))):
            result = finder.Execute(Replace=2 if replace_all else 1)
        count = int(result) if isinstance(result, (int, float)) else None
        return {"op": "word_live_replace_text", "find": find_text, "replace": replacement, "replace_all": replace_all, "replacements": count}

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
                paragraph.Range.ListFormat.ApplyListTemplateWithLevel(template, ContinuePreviousList=True, ApplyTo=0, DefaultListBehavior=0, Level=level)
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
        if args.get("alignment") is not None:
            table.Rows.Alignment = _enum(args["alignment"], {"left": 0, "center": 1, "right": 2}, "alignment")
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
        rng.InsertCrossReference(ReferenceType=reference_type, ReferenceItem=item, InsertAsHyperlink=bool(args.get("as_hyperlink", args.get("insert_as_hyperlink", True))))
        return {"op": "word_live_insert_cross_reference", "reference_type": reference_type, "reference_item": item}

    def _word_live_insert_equation(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        equation = _required_string(args, "equation")
        if args.get("paragraph_index"):
            rng = doc.Paragraphs(int(args["paragraph_index"])).Range
            _collapse_end(rng)
        else:
            rng = _insert_range(doc, args.get("position", "end"), args.get("bookmark"))
        rng.Text = equation
        math = doc.OMaths.Add(rng)
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
        start = doc.GoTo(What=1, Which=1, Count=page)
        end = doc.GoTo(What=1, Which=1, Count=end_page + 1)
        return {"op": "word_live_get_page_text", "page": page, "end_page": end_page, "text": str(_safe_value(doc.Range(start.Start, end.Start), "Text", ""))}

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
        comment = doc.Comments.Add(Range=rng, Text=args.get("comment_text", args.get("text", "")), Author=args.get("author", "white-collar"), Initials=args.get("initials", "WC"))
        return {"op": "word_live_add_comment", "comment": _comment_dict(comment, _count(doc.Comments))}

    def _word_live_list_revisions(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        revisions = []
        for index, revision in enumerate(_iter_collection(doc.Revisions), start=1):
            revisions.append({"id": index, "type": _safe_value(revision, "Type"), "author": _safe_value(revision, "Author"), "date": str(_safe_value(revision, "Date", "")), "text": str(_safe_value(revision.Range, "Text", ""))})
        return {"op": "word_live_list_revisions", "revision_count": len(revisions), "revisions": revisions}

    def _word_live_reply_to_comment(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        comment = _comment(doc, args)
        reply = comment.Replies.Add(Text=args.get("reply_text", args.get("text", "")), Author=args.get("author", "white-collar"), Initials=args.get("initials", "WC"))
        return {"op": "word_live_reply_to_comment", "comment": _comment_dict(reply, int(args.get("comment_index", 1)))}

    def _word_live_resolve_comment(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        comment = _comment(doc, args)
        if hasattr(comment, "Done"):
            comment.Done = bool(args.get("resolved", args.get("resolve", True)))
        else:
            comment.Resolved = bool(args.get("resolved", args.get("resolve", True)))
        return {"op": "word_live_resolve_comment", "resolved": bool(args.get("resolved", args.get("resolve", True)))}

    def _word_live_delete_comment(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        comment = _comment(doc, args)
        comment.Delete()
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
        target.PageNumbers.Add(PageNumberAlignment=_enum(args.get("alignment", "center"), {"left": 0, "center": 1, "right": 2}, "alignment"))
        if args.get("prefix"):
            target.Range.InsertBefore(args["prefix"])
        if args.get("include_total"):
            field_range = target.Range.Duplicate
            _collapse_end(field_range)
            field_range.InsertAfter(" / ")
            doc.Fields.Add(Range=field_range, Type=26)
        if args.get("suffix"):
            target.Range.InsertAfter(args["suffix"])
        return {"op": "word_live_add_page_numbers", "section": int(args.get("section_index", 1)), "position": position}

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

    def _word_live_undo(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        count = max(1, int(args.get("count", args.get("times", 1))))
        for _ in range(count):
            doc.Undo()
        return {"op": "word_live_undo", "count": count}

    def _word_screen_capture(self, app: Any, doc: Any, args: dict[str, Any]) -> dict[str, Any]:
        output = _absolute_output(_required_string(args, "output_path"))
        hwnd = int(doc.ActiveWindow.Hwnd)
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


def _default_screenshot(hwnd: int, output: Path) -> None:
    try:
        from PIL import ImageGrab
    except ImportError as exc:
        raise BackendUnavailableError("word-screen-capture") from exc
    image = ImageGrab.grab(window=hwnd)
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output, format="PNG")


def _find_document(app: Any, target: str) -> Any:
    wanted = Path(target).resolve()
    for doc in _iter_collection(getattr(app, "Documents", [])):
        candidates = {str(_safe_value(doc, "FullName", "")), str(_safe_value(doc, "Name", ""))}
        if str(wanted) in candidates or target in candidates or Path(str(_safe_value(doc, "FullName", ""))).resolve() == wanted:
            return doc
    raise ValidationError("target Word document is not open", details={"target": target})


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
def _tracking(doc: Any, enabled: bool) -> Iterator[None]:
    if not enabled:
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


def _style_name(paragraph: Any) -> str | None:
    style = _safe_value(paragraph, "Style")
    return str(_safe_value(style, "NameLocal", style)) if style is not None else None


def _range_info(rng: Any) -> dict[str, Any]:
    return {"start": _safe_value(rng, "Start"), "end": _safe_value(rng, "End")}


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


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _doc_key(doc: Any) -> str:
    return str(_safe_value(doc, "FullName", _safe_value(doc, "Name", id(doc))))
