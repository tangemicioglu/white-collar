from __future__ import annotations

import base64
import os
from pathlib import Path
import uuid
import zipfile

import pytest

from whitecollar.adapters.word_com import Win32WordComAdapter
from whitecollar.models import Plan
from whitecollar.word_ops import WORD_COM_MUTATING_OPERATIONS, WORD_COM_OPERATIONS, WORD_COM_READ_OPERATIONS


pytestmark = pytest.mark.real_word


def _plan(target: Path, operation: str, args: dict, root: Path, sequence: int) -> Plan:
    read = operation in WORD_COM_READ_OPERATIONS
    if read:
        policy = "read-only"
        write = {"mode": "none"}
    elif operation == "word_live_export_pdf":
        policy = "edit"
        write = {"mode": "save-as", "path": str(root / f"export-{sequence:03d}.pdf")}
    elif operation == "word_live_compare_documents":
        policy = "edit"
        write = {"mode": "save-as", "path": str(root / f"compare-{sequence:03d}.docx")}
    elif operation == "word_screen_capture":
        policy = "review"
        write = {"mode": "save-as", "path": str(root / f"screen-source-{sequence:03d}.docx")}
    else:
        policy = "edit"
        write = {"mode": "in-place", "snapshot": str(root / f"snapshot-{sequence:02d}-{operation}.docx")}
    return Plan.from_dict(
        {
            "schema": "white-collar.plan/v1",
            "app": "word",
            "target": {"path": str(target)},
            "policy": policy,
            "operations": [{"op": operation, "args": args}],
            "write": write,
        }
    )


def _new_document(app, path: Path):
    document = app.Documents.Add()
    document.Range(0, 0).InsertAfter("Draft target text\rHeading One\rHeading Two\r")
    document.Paragraphs(2).Style = document.Styles("Heading 1")
    document.Paragraphs(3).Style = document.Styles("Heading 2")
    document.SaveAs2(FileName=str(path), FileFormat=12)
    return document


def _write_png(path: Path) -> None:
    # A valid 1x1 PNG, kept inline so the integration test has no image dependency.
    path.write_bytes(base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
    ))


def _current_document(app, target: Path):
    wanted = target.resolve()
    for document in list(app.Documents):
        if Path(str(document.FullName)).resolve() == wanted:
            return document
    raise AssertionError(f"target is not open in Word: {target}")


def _document_text(document) -> str:
    return str(document.Content.Text)


def _assert_valid_word_copy(app, path: Path) -> None:
    assert path.is_file(), path
    assert zipfile.is_zipfile(path), path
    opened = app.Documents.Open(FileName=str(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        assert int(opened.Paragraphs.Count) > 0
    finally:
        opened.Close(SaveChanges=False)
    opened = None


def _assert_screenshot(path: Path) -> None:
    from PIL import Image, ImageStat

    assert path.is_file(), path
    with Image.open(path) as image:
        image.load()
        assert image.width > 100 and image.height > 100
        stats = ImageStat.Stat(image.convert("RGB"))
        assert max(stats.stddev) > 0, path


def _assert_pdf(path: Path) -> None:
    assert path.is_file(), path
    assert path.read_bytes()[:5] == b"%PDF-"


def _artifact_root(tmp_path: Path) -> Path:
    configured = os.environ.get("WHITE_COLLAR_REAL_WORD_ARTIFACT_DIR")
    if not configured:
        return tmp_path / "word-artifacts"
    root = Path(configured).resolve() / f"run-{uuid.uuid4().hex[:10]}"
    root.mkdir(parents=True, exist_ok=False)
    return root


def _assert_operation_behavior(app, target: Path, operation: str, args: dict, value: dict, before_text: str) -> None:
    result = value["operations"][0]
    document = _current_document(app, target)
    text = _document_text(document)

    if operation == "word_live_list_open":
        assert result["documents"]
    elif operation == "word_live_get_text":
        assert result["total_paragraphs"] >= 3
        assert any(item["text"].startswith("Draft") for item in result["paragraphs"])
    elif operation == "word_live_take_snapshot":
        assert result["paragraphs"] >= 3
    elif operation == "word_live_snapshot_status":
        assert result["exists"] is True
    elif operation == "word_live_get_page_text":
        assert "Draft" in result["text"]
    elif operation == "word_live_get_paragraph_format":
        assert result["paragraphs"]
    elif operation == "word_live_get_info":
        assert result["path"] == str(target)
        assert result["paragraphs"] >= 3
    elif operation == "word_live_find_text":
        assert any(match["text"] == "Draft" for match in result["matches"])
    elif operation == "word_live_list_cross_reference_items":
        assert any(item["text"] == "Heading One" for item in result["headings"])
    elif operation == "word_live_diagnose_layout":
        assert isinstance(result["warnings"], list)
    elif operation == "word_live_get_comments":
        if document.Comments.Count:
            assert result["comments"]
    elif operation == "word_live_list_styles":
        assert result["count"] >= 1
        assert any(item["name"] in {"Normal", "Normal (Web)"} for item in result["styles"])
    elif operation == "word_live_list_hyperlinks":
        assert "hyperlinks" in result
    elif operation == "word_live_list_notes":
        assert "notes" in result
    elif operation == "word_live_list_content_controls":
        assert "content_controls" in result
    elif operation == "word_live_get_protection":
        assert result["protected"] is False
    elif operation == "word_live_insert_text":
        assert "inserted" in text
    elif operation == "word_live_get_diff":
        assert result["changes"]
    elif operation == "word_live_insert_paragraphs":
        assert "Inserted A" in text and "Inserted B" in text
    elif operation == "word_live_format_text":
        font = document.Paragraphs(1).Range.Font
        assert bool(font.Bold) is True
        assert str(font.Name) == "Arial"
        assert abs(float(font.Size) - 12) < 0.1
    elif operation == "word_live_apply_style":
        assert str(document.Paragraphs(int(args.get("paragraph_index", 1))).Style.NameLocal) == str(args["style_name"])
    elif operation == "word_live_add_hyperlink":
        assert document.Hyperlinks.Count >= 1
    elif operation == "word_live_remove_hyperlink":
        assert document.Hyperlinks.Count == 0
    elif operation == "word_live_add_note":
        collection = document.Endnotes if args.get("note_type") == "endnote" else document.Footnotes
        assert collection.Count >= 1
    elif operation == "word_live_update_fields":
        assert result["fields"] >= 0
    elif operation == "word_live_insert_toc":
        assert document.TablesOfContents.Count >= 1
    elif operation == "word_live_set_content_control":
        assert document.ContentControls.Count >= 1
        assert any(
            str(document.ContentControls(index).Range.Text).rstrip("\r\a") == str(args["value"])
            for index in range(1, int(document.ContentControls.Count) + 1)
        )
    elif operation == "word_live_add_table":
        assert document.Tables.Count >= 1
        assert str(document.Tables(1).Cell(1, 1).Range.Text).startswith("A")
    elif operation == "word_live_format_table":
        assert int(document.Tables(1).Rows.Alignment) == 1
    elif operation == "word_live_apply_list":
        assert int(document.Paragraphs(1).Range.ListFormat.ListType) != 0
    elif operation == "word_live_setup_heading_numbering":
        assert int(document.Paragraphs(2).Range.ListFormat.ListLevelNumber) >= 1
    elif operation == "word_live_modify_table":
        assert str(document.Tables(1).Cell(1, 1).Range.Text).startswith("Changed")
    elif operation == "word_live_toggle_track_changes":
        assert bool(document.TrackRevisions) is True
    elif operation == "word_live_insert_image":
        assert document.InlineShapes.Count >= 1
        assert abs(float(document.InlineShapes(1).Width) - 16) < 0.5
        assert abs(float(document.InlineShapes(1).Height) - 16) < 0.5
    elif operation == "word_live_add_bookmark":
        assert document.Bookmarks.Exists("CrossRef")
    elif operation == "word_live_insert_cross_reference":
        assert document.Fields.Count >= 1
    elif operation == "word_live_insert_equation":
        assert document.OMaths.Count >= 1
    elif operation == "word_live_set_page_layout":
        page = document.Sections(1).PageSetup
        assert int(page.Orientation) == 0
        assert abs(float(page.PageWidth) - 612) < 1
        assert abs(float(page.PageHeight) - 792) < 1
    elif operation == "word_live_add_header_footer":
        assert "Header" in str(document.Sections(1).Headers(1).Range.Text)
        assert "Footer" in str(document.Sections(1).Footers(1).Range.Text)
    elif operation == "word_live_remove_header_footer":
        assert "Header" not in str(document.Sections(1).Headers(1).Range.Text)
        assert "Footer" not in str(document.Sections(1).Footers(1).Range.Text)
    elif operation == "word_live_add_watermark":
        assert document.Sections(1).Headers(1).Shapes.Count >= 1
    elif operation == "word_live_remove_watermark":
        remaining = 0
        for section in document.Sections:
            for kind in (1, 2, 3):
                for area in (section.Headers(kind), section.Footers(kind)):
                    for shape in area.Shapes:
                        try:
                            value = str(shape.TextEffect.Text or "").strip()
                        except Exception:
                            value = ""
                        remaining += value.casefold() == args.get("text", "DRAFT").casefold()
        assert remaining == 0
        assert result["removed"] >= 1
    elif operation == "word_live_add_page_numbers":
        container = document.Sections(1).Headers(1) if args.get("position") == "header" else document.Sections(1).Footers(1)
        assert container.Range.Fields.Count >= (2 if args.get("include_total") else 1)
    elif operation == "word_live_add_section_break":
        assert document.Sections.Count >= 2
    elif operation == "word_live_set_paragraph_spacing":
        assert abs(float(document.Paragraphs(1).Format.SpaceAfter) - 6) < 0.1
    elif operation == "word_live_add_comment":
        assert document.Comments.Count >= 1
        assert "Review this" in str(document.Comments(1).Range.Text)
    elif operation == "word_live_reply_to_comment":
        assert document.Comments(1).Replies.Count >= 1
        assert "Noted" in str(document.Comments(1).Replies(1).Range.Text)
    elif operation == "word_live_resolve_comment":
        with zipfile.ZipFile(target) as archive:
            assert b'w15:done="1"' in archive.read("word/commentsExtended.xml")
    elif operation == "word_live_delete_comment":
        assert document.Comments.Count == 0
    elif operation == "word_live_replace_text":
        assert args["replace_text"] in text
    elif operation == "word_live_list_revisions":
        assert result["revision_count"] > 0
    elif operation == "word_live_accept_revisions":
        assert document.Revisions.Count == 0
    elif operation == "word_live_reject_revisions":
        assert document.Revisions.Count == 0
    elif operation == "word_live_get_undo_history":
        assert result["entries"]
    elif operation == "word_live_save":
        assert bool(document.Saved) is True
    elif operation == "word_live_undo":
        assert result["count"] == 1
    elif operation == "word_live_set_core_properties":
        assert str(document.BuiltInDocumentProperties("Title").Value) == "Real Word Test"
        assert str(document.BuiltInDocumentProperties("Author").Value) == "white-collar"
    elif operation == "word_live_delete_text":
        assert result["characters"] == int(args["end"]) - int(args["start"])
        # With TrackRevisions enabled Word keeps the deleted characters in the
        # content stream; the revision itself is the observable mutation.
        assert len(text) < len(before_text) or document.Revisions.Count > 0
    elif operation == "word_live_export_pdf":
        _assert_pdf(Path(result["path"]))
    elif operation == "word_live_set_protection":
        expected = -1 if args["protection_type"] == "none" else {"tracked_changes": 0, "comments": 1, "forms": 2, "read_only": 3}[args["protection_type"]]
        assert int(document.ProtectionType) == expected
    elif operation == "word_live_compare_documents":
        _assert_valid_word_copy(app, Path(result["path"]))
    elif operation == "word_live_merge_document":
        assert "Merged source marker" in text


@pytest.fixture(scope="module")
def real_word():
    from win32com.client import DispatchEx

    app = DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        yield app
    finally:
        try:
            documents = list(app.Documents)
        except Exception:
            documents = []
        opened_document = None
        for opened_document in documents:
            try:
                opened_document.Close(SaveChanges=False)
            except Exception:
                pass
        opened_document = None
        documents = None
        try:
            app.Quit(SaveChanges=False)
        except Exception:
            pass


def test_every_registered_word_operation_against_real_word(real_word, tmp_path):
    document_path = tmp_path / "real-word.docx"
    document = _new_document(real_word, document_path)
    adapter = Win32WordComAdapter(app_factory=lambda: real_word)
    target = Path(str(document.FullName))
    artifact_root = _artifact_root(tmp_path)
    screenshot_root = artifact_root / "screenshots"
    screenshot_root.mkdir(parents=True, exist_ok=True)
    image_path = artifact_root / "pixel.png"
    _write_png(image_path)
    compare_source = artifact_root / "compare-source.docx"
    source_document = real_word.Documents.Add()
    source_document.Range(0, 0).InsertAfter("Compared source text\r")
    source_document.SaveAs2(FileName=str(compare_source), FileFormat=12)
    source_document.Close(SaveChanges=False)
    merge_source = artifact_root / "merge-source.docx"
    merge_document = real_word.Documents.Add()
    merge_document.Range(0, 0).InsertAfter("Merged source marker\r")
    merge_document.SaveAs2(FileName=str(merge_source), FileFormat=12)
    merge_document.Close(SaveChanges=False)

    # State needed by several operations is created by the adapter itself or by
    # the immediately preceding operation. Every operation is still invoked via
    # the public plan/apply boundary.
    cases = [
        ("word_live_list_open", {}),
        ("word_live_get_text", {}),
        ("word_live_take_snapshot", {}),
        ("word_live_snapshot_status", {}),
        ("word_live_get_page_text", {"page": 1}),
        ("word_live_get_paragraph_format", {"start_paragraph": 1, "end_paragraph": 2}),
        ("word_live_get_info", {}),
        ("word_live_list_styles", {}),
        ("word_live_list_hyperlinks", {}),
        ("word_live_list_notes", {}),
        ("word_live_list_content_controls", {}),
        ("word_live_get_protection", {}),
        ("word_live_find_text", {"search_text": "Draft", "max_results": 10}),
        ("word_live_list_cross_reference_items", {}),
        ("word_live_diagnose_layout", {}),
        ("word_live_get_comments", {}),
        ("word_live_insert_text", {"text": " inserted", "position": "end"}),
        ("word_live_get_diff", {}),
        ("word_live_insert_paragraphs", {"paragraphs": ["Inserted A", "Inserted B"], "position": "end"}),
        ("word_live_format_text", {"start_paragraph": 1, "end_paragraph": 1, "bold": True, "font_name": "Arial", "font_size": 12}),
        ("word_live_apply_style", {"paragraph_index": 1, "style_name": "Title"}),
        ("word_live_add_hyperlink", {"paragraph_index": 1, "url": "https://example.com", "display_text": "Example"}),
        ("word_live_list_hyperlinks", {}),
        ("word_live_remove_hyperlink", {"hyperlink_index": 1}),
        ("word_live_add_note", {"paragraph_index": 1, "text": "Footnote text", "note_type": "footnote"}),
        ("word_live_list_notes", {}),
        ("word_live_insert_toc", {"position": "start", "lower_heading_level": 2}),
        ("word_live_update_fields", {}),
        ("word_live_set_content_control", {"paragraph_index": 1, "title": "ClientName", "value": "Example Client", "tag": "client-name"}),
        ("word_live_list_content_controls", {}),
        ("word_live_add_table", {"rows": 2, "cols": 2, "data": [["A", "B"], ["C", "D"]]}),
        ("word_live_format_table", {"table_index": -1, "autofit": "window", "table_alignment": "center"}),
        ("word_live_apply_list", {"start_paragraph": 1, "end_paragraph": 2, "list_type": "bullet"}),
        ("word_live_setup_heading_numbering", {"h1_paragraphs": [2], "h2_paragraphs": [3]}),
        ("word_live_modify_table", {"table_index": 1, "operation": "set_cell", "row": 1, "col": 1, "text": "Changed"}),
        ("word_live_toggle_track_changes", {"enable": True}),
        ("word_live_insert_image", {"image_path": str(image_path), "position": "end", "width_pt": 16, "height_pt": 16}),
        ("word_live_add_bookmark", {"paragraph_index": 1, "bookmark_name": "CrossRef"}),
        ("word_live_list_cross_reference_items", {}),
        ("word_live_insert_cross_reference", {"ref_type": "Bookmark", "ref_item": "CrossRef", "paragraph_index": 1}),
        ("word_live_insert_equation", {"equation": "x = y", "position": "end"}),
        ("word_live_set_page_layout", {"orientation": "portrait", "page_width_inches": 8.5, "page_height_inches": 11}),
        ("word_live_add_header_footer", {"header_text": "Header", "footer_text": "Footer"}),
        ("word_live_add_watermark", {"text": "DRAFT"}),
        ("word_live_remove_watermark", {"text": "DRAFT"}),
        ("word_live_add_page_numbers", {"position": "footer", "alignment": "center", "prefix": "Page "}),
        ("word_live_add_page_numbers", {"position": "header", "alignment": "right", "prefix": "Page ", "include_total": True, "suffix": " total"}),
        ("word_live_remove_header_footer", {"position": "both", "section_index": 1}),
        ("word_live_add_section_break", {"break_type": "new_page"}),
        ("word_live_set_paragraph_spacing", {"paragraph_index": 1, "space_after_pt": 6, "line_spacing_rule": "single"}),
        ("word_live_add_comment", {"start": 0, "end": 5, "text": "Review this"}),
        ("word_live_get_comments", {}),
        ("word_live_reply_to_comment", {"comment_index": 1, "text": "Noted"}),
        ("word_live_resolve_comment", {"comment_index": 1, "resolve": True}),
        ("word_live_delete_comment", {"comment_index": 1}),
        ("word_live_replace_text", {"find_text": "Draft", "replace_text": "Tracked", "replace_all": True, "track_changes": True}),
        ("word_live_list_revisions", {}),
        ("word_live_accept_revisions", {}),
        ("word_live_replace_text", {"find_text": "Heading", "replace_text": "Rejected", "replace_all": True, "track_changes": True}),
        ("word_live_reject_revisions", {}),
        ("word_live_get_undo_history", {}),
        ("word_live_save", {}),
        ("word_live_undo", {"times": 1}),
        ("word_live_set_core_properties", {"title": "Real Word Test", "author": "white-collar"}),
        ("word_live_merge_document", {"source_path": str(merge_source)}),
        ("word_live_compare_documents", {"source_path": str(compare_source)}),
        ("word_live_export_pdf", {}),
        ("word_live_set_protection", {"protection_type": "read_only"}),
        ("word_live_get_protection", {}),
        ("word_live_set_protection", {"protection_type": "none"}),
        ("word_screen_capture", {"output_path": str(screenshot_root / "final-word-window.png")}),
        ("word_live_delete_text", {"start": 0, "end": 1}),
    ]

    executed = set()
    real_word.Visible = False
    for sequence, (operation, args) in enumerate(cases, start=1):
        before_text = _document_text(_current_document(real_word, target))
        if operation == "word_screen_capture":
            real_word.Visible = True
        try:
            value = adapter.apply(_plan(target, operation, args, artifact_root, sequence), dry_run=False)
        except Exception as exc:
            pytest.fail(f"real Word operation {operation} failed: {type(exc).__name__}: {exc}")
        finally:
            real_word.Visible = False
        assert value["backend"] == "word-com", operation
        assert value["operations"], operation
        _assert_operation_behavior(real_word, target, operation, args, value, before_text)
        executed.add(operation)

        if operation == "word_screen_capture":
            _assert_screenshot(Path(args["output_path"]))
            _assert_valid_word_copy(real_word, artifact_root / f"screen-source-{sequence:03d}.docx")
        elif operation in {"word_live_compare_documents", "word_live_export_pdf"}:
            if operation == "word_live_compare_documents":
                _assert_valid_word_copy(real_word, Path(value["operations"][0]["path"]))
            else:
                _assert_pdf(Path(value["operations"][0]["path"]))
        elif operation in WORD_COM_MUTATING_OPERATIONS and operation not in {
            "word_live_apply_style",
            "word_live_add_hyperlink",
            "word_live_remove_hyperlink",
            "word_live_add_note",
            "word_live_update_fields",
            "word_live_insert_toc",
            "word_live_set_content_control",
            "word_live_remove_header_footer",
            "word_live_set_protection",
            "word_live_merge_document",
        }:
            snapshot = artifact_root / f"snapshot-{sequence:02d}-{operation}.docx"
            _assert_valid_word_copy(real_word, snapshot)
            screenshot = screenshot_root / f"after-{sequence:02d}-{operation}.png"
            real_word.Visible = True
            try:
                capture = adapter.apply(
                    _plan(
                        target,
                        "word_screen_capture",
                        {"output_path": str(screenshot)},
                        artifact_root,
                        sequence * 1000,
                    ),
                    dry_run=False,
                )
            except Exception as exc:
                pytest.fail(f"screenshot after {operation} failed: {type(exc).__name__}: {exc}")
            finally:
                real_word.Visible = False
            assert capture["operations"][0]["op"] == "word_screen_capture"
            _assert_screenshot(screenshot)
            _assert_valid_word_copy(real_word, artifact_root / f"screen-source-{sequence * 1000:03d}.docx")

    # Creation has a separate disposable-target test because it does not act
    # on the shared open document used by the operation matrix below.
    assert WORD_COM_OPERATIONS - executed == {"word_live_create_document"}
    expected_screenshots = 1 + sum(
        operation in WORD_COM_MUTATING_OPERATIONS
        and operation not in {
            "word_live_compare_documents", "word_live_export_pdf",
            "word_live_apply_style", "word_live_add_hyperlink", "word_live_remove_hyperlink",
            "word_live_add_note", "word_live_update_fields", "word_live_insert_toc",
            "word_live_set_content_control", "word_live_remove_header_footer",
            "word_live_set_protection", "word_live_merge_document",
        }
        for operation, _ in cases
        if operation != "word_screen_capture"
    )
    assert len(list(screenshot_root.glob("*.png"))) == expected_screenshots
    document = None
    adapter = None


def test_remove_watermark_removes_header_and_footer_with_tracking(real_word, tmp_path):
    document_path = tmp_path / "watermark-header-footer.docx"
    document = _new_document(real_word, document_path)
    header_shape = document.Sections(1).Headers(1).Shapes.AddTextEffect(
        0, "DRAFT", "Calibri", 72, False, False, 0, 0
    )
    footer_shape = document.Sections(1).Footers(1).Shapes.AddTextEffect(
        0, "DRAFT", "Calibri", 72, False, False, 0, 0
    )
    header_shape = None
    footer_shape = None
    document.TrackRevisions = True
    document.Save()
    adapter = Win32WordComAdapter(app_factory=lambda: real_word)
    plan = _plan(
        document_path,
        "word_live_remove_watermark",
        {"text": "DRAFT", "position": "both"},
        tmp_path / "watermark-artifacts",
        1,
    )

    try:
        value = adapter.apply(plan, dry_run=False)
        current = _current_document(real_word, document_path)
        remaining = []
        for section in current.Sections:
            for kind in (1, 2, 3):
                for area in (section.Headers(kind), section.Footers(kind)):
                    for shape in area.Shapes:
                        try:
                            shape_text = str(shape.TextEffect.Text or "").strip()
                        except Exception:
                            shape_text = ""
                        if shape_text.casefold() == "draft":
                            remaining.append(shape_text)
        assert value["operations"][0]["removed"] >= 2
        assert remaining == []
        assert bool(current.TrackRevisions) is True
    finally:
        document.Close(SaveChanges=False)


def test_create_word_document_from_scratch_against_real_word(real_word, tmp_path):
    target = tmp_path / "created-from-scratch.docx"
    plan = Plan.from_dict(
        {
            "schema": "white-collar.plan/v1",
            "app": "word",
            "target": {"path": str(target)},
            "policy": "review",
            "operations": [{"op": "word_live_create_document"}],
            "write": {"mode": "create"},
        }
    )
    adapter = Win32WordComAdapter(app_factory=lambda: real_word)

    preview = adapter.apply(plan, dry_run=True)
    assert preview["written"] is False
    assert not target.exists()

    value = adapter.apply(plan, dry_run=False)
    assert value["backend"] == "word-com"
    assert value["operations"] == [{"op": "word_live_create_document", "created": True, "path": str(target)}]
    _assert_valid_word_copy(real_word, target)
