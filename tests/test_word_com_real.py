from __future__ import annotations

import base64
from pathlib import Path

import pytest

from whitecollar.adapters.word_com import Win32WordComAdapter
from whitecollar.models import Plan
from whitecollar.word_ops import WORD_COM_OPERATIONS, WORD_COM_READ_OPERATIONS


pytestmark = pytest.mark.real_word


def _plan(target: Path, operation: str, args: dict, root: Path, sequence: int) -> Plan:
    read = operation in WORD_COM_READ_OPERATIONS
    if read:
        policy = "read-only"
        write = {"mode": "none"}
    elif operation == "word_screen_capture":
        policy = "review"
        write = {"mode": "save-as", "path": str(root / "screen-source.docx")}
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
    document.Range(0, 0).InsertAfter("Draft target text\r")
    heading = document.Paragraphs.Add()
    heading.Range.Text = "Heading One"
    heading.Style = document.Styles("Heading 1")
    second = document.Paragraphs.Add()
    second.Range.Text = "Heading Two"
    second.Style = document.Styles("Heading 2")
    document.SaveAs2(FileName=str(path), FileFormat=12)
    return document


def _write_png(path: Path) -> None:
    # A valid 1x1 PNG, kept inline so the integration test has no image dependency.
    path.write_bytes(base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
    ))


@pytest.fixture(scope="module")
def real_word():
    from win32com.client import DispatchEx

    app = DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        yield app
    finally:
        documents = list(app.Documents)
        opened_document = None
        for opened_document in documents:
            try:
                opened_document.Close(SaveChanges=False)
            except Exception:
                pass
        opened_document = None
        documents = None
        app.Quit(SaveChanges=False)


def test_every_registered_word_operation_against_real_word(real_word, tmp_path):
    document_path = tmp_path / "real-word.docx"
    document = _new_document(real_word, document_path)
    adapter = Win32WordComAdapter(app_factory=lambda: real_word)
    target = Path(str(document.FullName))
    image_path = tmp_path / "pixel.png"
    _write_png(image_path)

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
        ("word_live_find_text", {"search_text": "Draft", "max_results": 10}),
        ("word_live_list_cross_reference_items", {}),
        ("word_live_diagnose_layout", {}),
        ("word_live_get_comments", {}),
        ("word_live_insert_text", {"text": " inserted", "position": "end"}),
        ("word_live_get_diff", {}),
        ("word_live_insert_paragraphs", {"paragraphs": ["Inserted A", "Inserted B"], "position": "end"}),
        ("word_live_format_text", {"start_paragraph": 1, "end_paragraph": 1, "bold": True, "font_name": "Arial", "font_size": 12}),
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
        ("word_live_add_page_numbers", {"position": "footer", "alignment": "center", "prefix": "Page "}),
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
        ("word_screen_capture", {"output_path": str(tmp_path / "word-window.png")}),
        ("word_live_delete_text", {"start": 0, "end": 1}),
    ]

    executed = set()
    real_word.Visible = False
    for sequence, (operation, args) in enumerate(cases, start=1):
        if operation == "word_screen_capture":
            real_word.Visible = True
        try:
            value = adapter.apply(_plan(target, operation, args, tmp_path, sequence), dry_run=False)
        except Exception as exc:
            pytest.fail(f"real Word operation {operation} failed: {type(exc).__name__}: {exc}")
        finally:
            real_word.Visible = False
        assert value["backend"] == "word-com", operation
        assert value["operations"], operation
        executed.add(operation)

    assert WORD_COM_OPERATIONS - executed == set()
    assert (tmp_path / "word-window.png").is_file()
    document = None
    adapter = None
