"""Record a real Word/PowerPoint showcase for the white-collar README.

The fixture files are disposable.  Direct COM is used only to seed the
fixtures and keep their windows visible; every operation in the showcase is
invoked through ``white-collar ... apply --plan``.
"""

from __future__ import annotations

import argparse
import ctypes
import gc
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from typing import Any, Callable

from whitecollar.slides_ops import SLIDES_COM_READ_OPERATIONS
from whitecollar.word_ops import WORD_COM_READ_OPERATIONS


ROOT = Path(__file__).resolve().parents[1]


def _window_for_fragment(fragment: str) -> tuple[int, str]:
    import win32gui

    matches: list[tuple[int, str]] = []

    def collect(hwnd: int, _extra: object) -> None:
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if fragment.lower() in title.lower():
                matches.append((int(hwnd), title))

    win32gui.EnumWindows(collect, None)
    if len(matches) != 1:
        raise RuntimeError(f"expected one visible Office window containing {fragment!r}, found {matches!r}")
    return matches[0]


def _maximize_window(hwnd: int) -> str:
    import win32gui

    win32gui.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    return win32gui.GetWindowText(hwnd)


def _ffmpeg_path() -> str:
    value = shutil.which("ffmpeg")
    if not value:
        raise RuntimeError("ffmpeg is required to record the demo; install it and put it on PATH")
    return value


def _window_image(hwnd: int, size: tuple[int, int] = (1280, 720)) -> Any:
    """Render one Office HWND without capturing any other desktop window."""

    from PIL import Image
    import win32gui
    import win32ui

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        raise RuntimeError("Office window has no visible dimensions")
    window_dc = win32gui.GetWindowDC(hwnd)
    source_dc = win32ui.CreateDCFromHandle(window_dc)
    memory_dc = source_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source_dc, width, height)
    memory_dc.SelectObject(bitmap)
    try:
        rendered = ctypes.windll.user32.PrintWindow(hwnd, memory_dc.GetSafeHdc(), 2)
        if not rendered:
            raise RuntimeError("PrintWindow could not render the Office window")
        bits = bitmap.GetBitmapBits(True)
        image = Image.frombuffer("RGB", (width, height), bits, "raw", "BGRX", 0, 1).copy()
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memory_dc.DeleteDC()
        source_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, window_dc)
    resampling = getattr(Image, "Resampling", Image).LANCZOS
    return image.resize(size, resampling)


def _capture_segment(ffmpeg: str, hwnd: int, output: Path, action: Callable[[], None]) -> None:
    """Capture one Office HWND until the showcase action finishes."""

    output.parent.mkdir(parents=True, exist_ok=True)
    command = [
        ffmpeg,
        "-y",
        "-hide_banner",
        "-loglevel",
        "error",
        "-f",
        "rawvideo",
        "-framerate",
        "7",
        "-pix_fmt",
        "rgb24",
        "-video_size",
        "1280x720",
        "-i",
        "-",
        "-r",
        "30",
        "-an",
        "-c:v",
        "libx264",
        "-preset",
        "veryfast",
        "-crf",
        "27",
        "-pix_fmt",
        "yuv420p",
        "-movflags",
        "+faststart",
        str(output),
    ]
    recorder = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    stop = threading.Event()
    capture_errors: list[BaseException] = []

    def capture_loop() -> None:
        try:
            while not stop.is_set():
                image = _window_image(hwnd)
                if recorder.stdin is None:
                    return
                recorder.stdin.write(image.tobytes())
                recorder.stdin.flush()
                time.sleep(1 / 15)
        except BaseException as exc:
            capture_errors.append(exc)
            stop.set()

    thread = threading.Thread(target=capture_loop, name="white-collar-window-capture", daemon=True)
    thread.start()
    try:
        action()
        time.sleep(3)
    finally:
        stop.set()
        thread.join(timeout=10)
        if recorder.stdin is not None:
            recorder.stdin.close()
        try:
            recorder.wait(timeout=30)
        except subprocess.TimeoutExpired:
            recorder.terminate()
            recorder.wait(timeout=5)
    if capture_errors:
        raise RuntimeError(f"Office window capture failed: {capture_errors[0]}")
    if recorder.returncode != 0 or not output.is_file():
        raise RuntimeError(f"ffmpeg failed to capture Office window into {output}")


def _run_cli(arguments: list[str]) -> dict[str, Any]:
    completed = subprocess.run(
        [sys.executable, "-m", "whitecollar.cli", *arguments],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )
    if completed.returncode:
        raise RuntimeError(
            f"CLI command failed with exit code {completed.returncode}: {' '.join(arguments)}\n"
            f"{completed.stdout}\n{completed.stderr}"
        )
    try:
        return json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"CLI did not return JSON for {' '.join(arguments)}: {completed.stdout!r}") from exc


def _write_plan(path: Path, plan: dict[str, Any]) -> None:
    path.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")


def _write_word_image(path: Path) -> None:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (180, 90), (31, 78, 121))
    draw = ImageDraw.Draw(image)
    draw.rectangle((12, 12, 168, 78), outline=(255, 255, 255), width=3)
    draw.text((30, 35), "WC", fill=(255, 255, 255))
    image.save(path)


def _write_slide_assets(ffmpeg: str, root: Path) -> dict[str, Path]:
    from PIL import Image, ImageDraw
    import wave

    image_path = root / "slide-image.png"
    image = Image.new("RGB", (160, 100), (32, 128, 224))
    ImageDraw.Draw(image).rectangle((10, 10, 150, 90), outline=(255, 255, 255), width=4)
    image.save(image_path)

    audio_path = root / "tone.wav"
    with wave.open(str(audio_path), "wb") as audio:
        audio.setnchannels(1)
        audio.setsampwidth(2)
        audio.setframerate(8000)
        audio.writeframes(b"\x00\x00" * 800)

    video_path = root / "tiny-video.mp4"
    completed = subprocess.run(
        [
            ffmpeg,
            "-y",
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "lavfi",
            "-i",
            "color=c=204060:s=160x100:d=1",
            "-pix_fmt",
            "yuv420p",
            "-c:v",
            "libx264",
            str(video_path),
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    if completed.returncode or not video_path.is_file():
        raise RuntimeError(f"could not create the disposable video asset: {completed.stderr}")
    return {"image": image_path, "audio": audio_path, "video": video_path}


def _word_fixture(source: Path, root: Path) -> tuple[Any, Any, dict[str, Path]]:
    from win32com.client import DispatchEx, GetActiveObject
    import win32process

    assets = {"image": root / "word-image.png"}
    _write_word_image(assets["image"])

    existing_pid = None
    try:
        active = GetActiveObject("Word.Application")
        existing_pid = int(win32process.GetWindowThreadProcessId(int(active.Windows(1).Hwnd))[1])
    except Exception:
        pass

    app = DispatchEx("Word.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    document = app.Documents.Add()
    document.Range(0, 0).InsertAfter(
        "Draft target text\rHeading One\rHeading Two\r"
        "A short paragraph for the white-collar showcase.\r"
    )
    document.Paragraphs(2).Style = document.Styles("Heading 1")
    document.Paragraphs(3).Style = document.Styles("Heading 2")
    document.SaveAs2(FileName=str(source), FileFormat=12)
    # Save one real comment into the fixture first.  Word then carries its
    # modern comments extension into save-as copies, which lets the CLI demo
    # exercise resolve_comment without requiring an in-place edit grant.
    document.Comments.Add(document.Range(0, 5), "Seed comment")
    document.Save()
    hwnd = int(app.Windows(1).Hwnd)
    pid = int(win32process.GetWindowThreadProcessId(hwnd)[1])
    if existing_pid is not None and pid == existing_pid:
        document.Close(SaveChanges=False)
        app.Quit(SaveChanges=False)
        raise RuntimeError("Word DispatchEx reused the user Word process; refusing to touch another window")

    compare_source = root / "compare-source.docx"
    compare_document = app.Documents.Add()
    compare_document.Range(0, 0).InsertAfter("Compared source text\r")
    compare_document.SaveAs2(FileName=str(compare_source), FileFormat=12)
    compare_document.Close(SaveChanges=False)
    merge_source = root / "merge-source.docx"
    merge_document = app.Documents.Add()
    merge_document.Range(0, 0).InsertAfter("Merged source marker\r")
    merge_document.SaveAs2(FileName=str(merge_source), FileFormat=12)
    merge_document.Close(SaveChanges=False)
    assets.update({"compare": compare_source, "merge": merge_source})
    document.Activate()
    _maximize_window(hwnd)
    return app, document, assets


def _slides_fixture(source: Path, root: Path, ffmpeg: str) -> tuple[Any, Any, dict[str, Path]]:
    from win32com.client import DispatchEx

    app = DispatchEx("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = 0
    presentation = app.Presentations.Add(True)
    for index, (title_text, body_text) in enumerate(
        (("Draft Review", "Draft status\rAgenda"), ("Second Slide", "Planning notes")),
        start=1,
    ):
        slide = presentation.Slides.Add(index, 12)  # ppLayoutBlank
        title = slide.Shapes.AddTextbox(1, 55, 40, 850, 70)
        title.Name = "Title"
        title.TextFrame.TextRange.Text = title_text
        title.TextFrame.TextRange.Font.Size = 28
        title.TextFrame.TextRange.Font.Bold = True
        body = slide.Shapes.AddTextbox(1, 55, 145, 850, 80)
        body.Name = "Body"
        body.TextFrame.TextRange.Text = body_text
        body.TextFrame.TextRange.Font.Size = 22
    presentation.SaveAs(FileName=str(source), FileFormat=24)
    assets = _write_slide_assets(ffmpeg, root)
    template = app.Presentations.Add(True)
    template.Slides.Add(1, 12)
    template.SaveAs(FileName=str(root / "source.potx"), FileFormat=26)
    template.Close()
    assets["template"] = root / "source.potx"
    hwnd, _title = _window_for_fragment(source.stem)
    _maximize_window(hwnd)
    return app, presentation, assets


def _close_word(app: Any, document: Any) -> None:
    try:
        document.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        app.Quit(SaveChanges=False)
    except Exception:
        pass


def _close_slides(app: Any, presentation: Any) -> None:
    try:
        presentation.Close()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass


def _output_for(app: str, operation: str, root: Path, sequence: int) -> Path:
    if operation in {"word_live_export_pdf", "slides_live_export_pdf"}:
        return root / "outputs" / f"{app}-{sequence:03d}.pdf"
    if operation == "slides_live_save_template":
        return root / "outputs" / f"slides-{sequence:03d}.potx"
    if operation == "word_screen_capture":
        return root / "outputs" / f"word-{sequence:03d}.docx"
    if operation == "slides_screen_capture":
        return root / "outputs" / f"slides-{sequence:03d}.pptx"
    return root / "outputs" / f"{app}-{sequence:03d}.{'docx' if app == 'word' else 'pptx'}"


def _apply_operation(
    app: str,
    target: Path,
    operation: str,
    args: dict[str, Any],
    root: Path,
    sequence: int,
    *,
    pause: float = 0.8,
) -> dict[str, Any]:
    read_operations = WORD_COM_READ_OPERATIONS if app == "word" else SLIDES_COM_READ_OPERATIONS
    if operation in read_operations:
        policy = "read-only"
        write = {"mode": "none"}
    else:
        policy = "review"
        output = _output_for(app, operation, root, sequence)
        output.parent.mkdir(parents=True, exist_ok=True)
        write = {"mode": "save-as", "path": str(output.resolve())}
    plan_path = root / "plans" / f"{app}-{sequence:03d}-{operation}.json"
    plan_path.parent.mkdir(parents=True, exist_ok=True)
    _write_plan(
        plan_path,
        {
            "schema": "white-collar.plan/v1",
            "app": app,
            "target": {"path": str(target.resolve())},
            "policy": policy,
            "operations": [{"op": operation, "args": args}],
            "write": write,
        },
    )
    value = _run_cli([app, "apply", "--plan", str(plan_path), "--backend", "com"])
    time.sleep(pause)
    return value


def _apply_batch(
    app: str,
    target: Path,
    operations: list[tuple[str, dict[str, Any]]],
    root: Path,
    sequence: int,
) -> dict[str, Any]:
    """Run stateful operations in one CLI process (snapshot/diff/undo)."""

    read_operations = WORD_COM_READ_OPERATIONS if app == "word" else SLIDES_COM_READ_OPERATIONS
    is_mutation = any(operation not in read_operations for operation, _args in operations)
    policy = "review" if is_mutation else "read-only"
    write = {"mode": "none"}
    if is_mutation:
        output = _output_for(app, operations[0][0], root, sequence)
        output.parent.mkdir(parents=True, exist_ok=True)
        write = {"mode": "save-as", "path": str(output.resolve())}
    plan_path = root / "plans" / f"{app}-{sequence:03d}-batch.json"
    plan_path.parent.mkdir(parents=True, exist_ok=True)
    _write_plan(
        plan_path,
        {
            "schema": "white-collar.plan/v1",
            "app": app,
            "target": {"path": str(target.resolve())},
            "policy": policy,
            "operations": [{"op": operation, "args": args} for operation, args in operations],
            "write": write,
        },
    )
    value = _run_cli([app, "apply", "--plan", str(plan_path), "--backend", "com"])
    time.sleep(1.2)
    return value


def _word_read_showcase(target: Path, root: Path) -> None:
    _apply_batch(
        "word",
        target,
        [
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
        ],
        root,
        1,
    )


def _word_mutation_showcase(target: Path, root: Path, assets: dict[str, Path]) -> None:
    _apply_batch(
        "word",
        target,
        [
            ("word_live_take_snapshot", {}),
            ("word_live_insert_text", {"text": " inserted", "position": "end"}),
            ("word_live_get_diff", {}),
        ],
        root,
        2,
    )
    _apply_batch(
        "word",
        target,
        [
            ("word_live_add_hyperlink", {"paragraph_index": 1, "url": "https://example.com", "display_text": "Example"}),
            ("word_live_list_hyperlinks", {}),
            ("word_live_remove_hyperlink", {"hyperlink_index": 1}),
            ("word_live_list_hyperlinks", {}),
        ],
        root,
        3,
    )
    _apply_batch(
        "word",
        target,
        [
            ("word_live_add_note", {"paragraph_index": 1, "text": "Footnote text", "note_type": "footnote"}),
            ("word_live_list_notes", {}),
        ],
        root,
        4,
    )
    _apply_batch(
        "word",
        target,
        [
            ("word_live_set_content_control", {"target_text": "Draft target text", "title": "ClientName", "value": "Example Client", "tag": "client-name"}),
            ("word_live_list_content_controls", {}),
        ],
        root,
        5,
    )
    _apply_batch(
        "word",
        target,
        [
            ("word_live_add_table", {"rows": 2, "cols": 2, "data": [["A", "B"], ["C", "D"]]}),
            ("word_live_format_table", {"table_index": -1, "autofit": "window", "table_alignment": "center"}),
            ("word_live_modify_table", {"table_index": 1, "operation": "set_cell", "row": 1, "col": 1, "text": "Changed"}),
        ],
        root,
        6,
    )
    _apply_batch(
        "word",
        target,
        [
            ("word_live_add_bookmark", {"paragraph_index": 1, "bookmark_name": "CrossRef"}),
            ("word_live_list_cross_reference_items", {}),
            ("word_live_insert_cross_reference", {"ref_type": "Bookmark", "ref_item": "CrossRef", "paragraph_index": 1}),
        ],
        root,
        7,
    )
    _apply_batch(
        "word",
        target,
        [
            ("word_live_add_comment", {"start": 6, "end": 11, "text": "Review this"}),
            ("word_live_get_comments", {}),
            ("word_live_reply_to_comment", {"comment_index": 2, "text": "Noted"}),
            ("word_live_resolve_comment", {"comment_index": 2, "resolve": True}),
            ("word_live_save", {}),
        ],
        root,
        8,
    )
    _apply_operation("word", target, "word_live_delete_comment", {"comment_index": 2}, root, 9)
    _apply_batch(
        "word",
        target,
        [
            ("word_live_replace_text", {"find_text": "Heading One", "replace_text": "Tracked", "replace_all": True, "track_changes": True}),
            ("word_live_list_revisions", {}),
            ("word_live_accept_revisions", {}),
            ("word_live_replace_text", {"find_text": "Heading", "replace_text": "Rejected", "replace_all": True, "track_changes": True}),
            ("word_live_reject_revisions", {}),
        ],
        root,
        10,
    )
    _apply_batch(
        "word",
        target,
        [
            ("word_live_set_protection", {"protection_type": "read_only"}),
            ("word_live_get_protection", {}),
            ("word_live_set_protection", {"protection_type": "none"}),
        ],
        root,
        11,
    )
    cases = [
        ("word_live_insert_paragraphs", {"paragraphs": ["Inserted A", "Inserted B"], "position": "end"}),
        ("word_live_format_text", {"start_paragraph": 1, "end_paragraph": 1, "bold": True, "font_name": "Arial", "font_size": 16}),
        ("word_live_apply_style", {"paragraph_index": 1, "style_name": "Title"}),
        ("word_live_insert_toc", {"position": "start", "lower_heading_level": 2}),
        ("word_live_update_fields", {}),
        ("word_live_apply_list", {"start_paragraph": 1, "end_paragraph": 2, "list_type": "bullet"}),
        ("word_live_setup_heading_numbering", {"h1_paragraphs": [2], "h2_paragraphs": [3]}),
        ("word_live_toggle_track_changes", {"enable": True}),
        ("word_live_insert_image", {"image_path": str(assets["image"]), "position": "end", "width_pt": 72, "height_pt": 40}),
        ("word_live_insert_equation", {"equation": "x = y", "position": "end"}),
        ("word_live_set_page_layout", {"orientation": "portrait", "page_width_inches": 8.5, "page_height_inches": 11}),
        ("word_live_add_header_footer", {"header_text": "White-collar showcase", "footer_text": "Live COM"}),
        ("word_live_add_watermark", {"text": "DRAFT"}),
        ("word_live_remove_watermark", {"text": "DRAFT"}),
        ("word_live_add_page_numbers", {"position": "footer", "alignment": "center", "prefix": "Page "}),
        ("word_live_add_page_numbers", {"position": "header", "alignment": "right", "prefix": "Page ", "include_total": True, "suffix": " total"}),
        ("word_live_remove_header_footer", {"position": "both", "section_index": 1}),
        ("word_live_add_section_break", {"break_type": "new_page"}),
        ("word_live_set_paragraph_spacing", {"paragraph_index": 1, "space_after_pt": 6, "line_spacing_rule": "single"}),
        ("word_live_get_undo_history", {}),
        ("word_live_save", {}),
        ("word_live_undo", {"times": 1}),
        ("word_live_set_core_properties", {"title": "White-collar showcase", "author": "white-collar"}),
        ("word_live_compare_documents", {"source_path": str(assets["compare"])}),
        ("word_live_export_pdf", {}),
    ]
    for sequence, (operation, args) in enumerate(cases, start=12):
        _apply_operation("word", target, operation, args, root, sequence)
    _apply_batch(
        "word",
        target,
        [
            ("word_live_merge_document", {"source_path": str(assets["merge"])}),
            ("word_live_delete_text", {"target_text": "Merged source marker"}),
        ],
        root,
        40,
    )


def _slides_read_showcase(target: Path, root: Path) -> None:
    _apply_batch(
        "slides",
        target,
        [
            ("slides_live_list_open", {}),
            ("slides_live_get_info", {}),
            ("slides_live_get_text", {}),
            ("slides_live_get_slide_text", {"slide_index": 1}),
            ("slides_live_find_text", {"search_text": "Draft"}),
            ("slides_live_get_masters", {}),
            ("slides_live_get_layouts", {}),
            ("slides_live_get_placeholders", {"master": 1}),
            ("slides_live_get_notes", {"slide_index": 1}),
            ("slides_live_get_sections", {}),
            ("slides_live_get_media", {}),
        ],
        root,
        1,
    )


def _slides_mutation_showcase(target: Path, root: Path, assets: dict[str, Path]) -> None:
    cases: list[tuple[str, dict[str, Any]]] = [
        ("slides_live_insert_text", {"slide_index": 1, "shape_name": "Body", "text": " inserted"}),
        ("slides_live_replace_text", {"find_text": "Draft", "replace_text": "Final", "replace_all": True}),
        ("slides_live_apply_template", {"source_path": str(assets["template"])}),
        ("slides_live_save_template", {}),
        ("slides_live_add_slide", {"slide_index": 3, "title": "New Slide"}),
        ("slides_live_set_title", {"slide_index": 3, "title": "Reviewed Slide"}),
        ("slides_live_add_textbox", {"slide_index": 3, "name": "Inserted Text", "text": "Added body", "top": 130}),
        ("slides_live_format_text", {"slide_index": 3, "shape_name": "Inserted Text", "font_name": "Arial", "font_size": 24, "bold": True}),
        ("slides_live_set_layout", {"slide_index": 3, "layout": 1}),
        ("slides_live_add_shape", {"slide_index": 3, "name": "Accent", "shape_type": "rectangle", "fill_color": "204060"}),
        ("slides_live_add_image", {"slide_index": 3, "name": "Logo", "image_path": str(assets["image"]), "top": 240, "width": 120, "height": 75}),
        ("slides_live_set_background", {"slide_index": 3, "color": "F0F4F8"}),
        ("slides_live_duplicate_slide", {"slide_index": 3}),
        ("slides_live_reorder_slide", {"slide_index": 4, "to_index": 2}),
        ("slides_live_set_notes", {"slide_index": 4, "text": "Review notes from the CLI"}),
        ("slides_live_get_notes", {"slide_index": 4}),
        ("slides_live_set_slide_size", {"width_inches": 10, "height_inches": 5.625}),
        ("slides_live_add_shape", {"slide_index": 4, "name": "GroupA", "shape_type": "rectangle", "left": 100, "top": 350}),
        ("slides_live_add_shape", {"slide_index": 4, "name": "GroupB", "shape_type": "rectangle", "left": 300, "top": 350}),
    ]
    sequence = 2
    for operation, args in cases:
        _apply_operation("slides", target, operation, args, root, sequence)
        sequence += 1

    group_result = _apply_operation(
        "slides", target, "slides_live_group", {"slide_index": 4, "shape_names": ["GroupA", "GroupB"]}, root, sequence
    )
    sequence += 1
    group_name = group_result["data"]["operations"][0]["shape"]
    _apply_operation("slides", target, "slides_live_ungroup", {"slide_index": 4, "shape_name": group_name}, root, sequence)
    sequence += 1

    cases = [
        ("slides_live_align", {"slide_index": 4, "shape_names": ["Accent", "Logo"], "alignment": "left"}),
        ("slides_live_distribute", {"slide_index": 4, "shape_names": ["Accent", "Logo", "Inserted Text"], "direction": "horizontal"}),
        ("slides_live_z_order", {"slide_index": 4, "shape_name": "Accent", "command": "bring_to_front"}),
        ("slides_live_crop_image", {"slide_index": 4, "shape_name": "Logo", "left": 2, "top": 2, "right": 1, "bottom": 1}),
        ("slides_live_rotate_shape", {"slide_index": 4, "shape_name": "Accent", "degrees": 15}),
        ("slides_live_add_section", {"name": "Review"}),
        ("slides_live_get_sections", {}),
        ("slides_live_delete_section", {"section_index": 1}),
        ("slides_live_set_slide_visibility", {"slide_index": 4, "visible": False}),
        ("slides_live_set_slide_numbers", {"visible": True}),
        ("slides_live_add_table", {"slide_index": 4, "name": "DataTable", "rows": 2, "columns": 2, "data": [["A", "B"], ["C", "D"]]}),
        ("slides_live_set_table_cell", {"slide_index": 4, "shape_name": "DataTable", "row": 1, "column": 1, "text": "Updated"}),
        ("slides_live_add_chart", {"slide_index": 4, "name": "DataChart", "chart_type": "column", "title": "Results", "data": [["Category", "Value"], ["A", 1], ["B", 2]]}),
        ("slides_live_add_smartart", {"slide_index": 4, "name": "FlowSmartArt", "layout": 1, "nodes": ["Start"]}),
        ("slides_live_add_media", {"slide_index": 4, "name": "Audio", "media_path": str(assets["audio"])}),
        ("slides_live_add_media", {"slide_index": 4, "name": "Video", "media_path": str(assets["video"])}),
        ("slides_live_get_media", {"slide_index": 4}),
        ("slides_live_set_hyperlink", {"slide_index": 4, "shape_name": "Accent", "url": "https://example.com"}),
        ("slides_live_set_alt_text", {"slide_index": 4, "shape_name": "Logo", "text": "Blue logo"}),
        ("slides_live_set_transition", {"slide_index": 4, "effect": "fade", "advance_on_click": True}),
        ("slides_live_add_animation", {"slide_index": 4, "shape_name": "Accent", "effect": "fade"}),
        ("slides_live_export_pdf", {}),
        ("slides_live_save", {}),
        ("slides_live_delete_slide", {"slide_index": 2}),
    ]
    for operation, args in cases:
        _apply_operation("slides", target, operation, args, root, sequence)
        sequence += 1


def record(output: Path, *, force: bool = False) -> Path:
    if output.exists() and not force:
        raise RuntimeError(f"output already exists; pass --force to replace it: {output}")
    ffmpeg = _ffmpeg_path()
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="white-collar-office-showcase-", ignore_cleanup_errors=True) as temporary_name:
        root = Path(temporary_name)

        word_source = root / "white-collar-word-showcase.docx"
        word_app, word_document, word_assets = _word_fixture(word_source, root)
        try:
            word_hwnd, _word_title = _window_for_fragment(word_source.stem)
            _maximize_window(int(word_app.Windows(1).Hwnd))
            _run_cli(["word", "inspect", str(word_source), "--backend", "com"])
            word_clip = root / "word-showcase.mp4"
            _capture_segment(
                ffmpeg,
                word_hwnd,
                word_clip,
                lambda: (_word_read_showcase(word_source, root), _word_mutation_showcase(word_source, root, word_assets)),
            )
        finally:
            _close_word(word_app, word_document)
            word_document = None
            word_app = None
            word_assets = None
            gc.collect()

        slides_source = root / "white-collar-slides-showcase.pptx"
        slides_app, slides_presentation, slides_assets = _slides_fixture(slides_source, root, ffmpeg)
        try:
            slides_hwnd, _slides_title = _window_for_fragment(slides_source.stem)
            _maximize_window(slides_hwnd)
            _run_cli(["slides", "inspect", str(slides_source), "--backend", "com"])
            slides_clip = root / "slides-showcase.mp4"
            _capture_segment(
                ffmpeg,
                slides_hwnd,
                slides_clip,
                lambda: (_slides_read_showcase(slides_source, root), _slides_mutation_showcase(slides_source, root, slides_assets)),
            )
        finally:
            _close_slides(slides_app, slides_presentation)
            slides_presentation = None
            slides_app = None
            slides_assets = None
            gc.collect()

        _concat(ffmpeg, [word_clip, slides_clip], output, root)

    print(f"recorded {output} ({output.stat().st_size:,} bytes)")
    return output


def _concat(ffmpeg: str, clips: list[Path], output: Path, temporary: Path) -> None:
    listing = temporary / "concat.txt"
    listing.write_text("\n".join(f"file '{clip.as_posix()}'" for clip in clips) + "\n", encoding="utf-8")
    completed = subprocess.run(
        [
            ffmpeg,
            "-y",
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(listing),
            "-c",
            "copy",
            "-movflags",
            "+faststart",
            str(output),
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    if completed.returncode or not output.is_file():
        raise RuntimeError(f"ffmpeg could not combine the showcase clips: {completed.stderr.strip()}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Record the full real-Office white-collar showcase")
    parser.add_argument("--output", default=str(ROOT / "docs" / "demo" / "white-collar-office-demo.mp4"))
    parser.add_argument("--force", action="store_true", help="replace an existing output video")
    args = parser.parse_args()
    record(Path(args.output), force=args.force)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
