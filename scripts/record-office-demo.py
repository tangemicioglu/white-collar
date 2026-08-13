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


def _show_slide(app: Any, presentation: Any, slide_index: int) -> None:
    """Put the slide being demonstrated in the visible PowerPoint editor."""

    try:
        app.Visible = True
    except Exception:
        pass
    try:
        window = presentation.Windows(1)
        window.Activate()
        window.View.GotoSlide(int(slide_index))
    except Exception:
        try:
            presentation.Slides(int(slide_index)).Select()
        except Exception:
            pass
    time.sleep(1.0)


def _activate_word_result(app: Any, output: Path, root: Path) -> Any:
    """Open the saved result for the next staged CLI plan.

    Review plans intentionally use save-as.  On Word builds where SaveCopyAs
    is unavailable, the adapter restores the original target after saving.  A
    showcase needs the saved result to become the next target so the visible
    document accumulates the staged changes.
    """

    output = output.resolve()
    matching = None
    documents = [app.Documents(index) for index in range(1, int(app.Documents.Count) + 1)]
    for document in documents:
        full_name = str(getattr(document, "FullName", ""))
        try:
            in_fixture = Path(full_name).resolve().is_relative_to(root.resolve())
        except (OSError, ValueError):
            in_fixture = False
        if not in_fixture:
            continue
        if Path(full_name).resolve() == output:
            matching = document
        else:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
    if matching is None:
        last_error = None
        for attempt in range(12):
            try:
                matching = app.Documents.Open(FileName=str(output), ReadOnly=False, AddToRecentFiles=False)
                break
            except Exception as exc:
                last_error = exc
                time.sleep(0.5 + attempt * 0.25)
        if matching is None:
            raise RuntimeError(f"Word did not reopen the staged result: {output}") from last_error
    matching.Activate()
    _maximize_window(int(app.Windows(1).Hwnd))
    return matching


def _activate_slides_result(app: Any, output: Path, root: Path) -> Any:
    """Open a staged PowerPoint result and make its editor window active."""

    output = output.resolve()
    matching = None
    presentations = [app.Presentations(index) for index in range(1, int(app.Presentations.Count) + 1)]
    for presentation in presentations:
        full_name = str(getattr(presentation, "FullName", ""))
        try:
            in_fixture = Path(full_name).resolve().is_relative_to(root.resolve())
        except (OSError, ValueError):
            in_fixture = False
        if not in_fixture:
            continue
        if Path(full_name).resolve() == output:
            matching = presentation
        else:
            try:
                presentation.Close()
            except Exception:
                pass
    if matching is None:
        last_error = None
        for attempt in range(16):
            try:
                matching = app.Presentations.Open(
                    FileName=str(output),
                    ReadOnly=False,
                    Untitled=False,
                    WithWindow=True,
                )
                break
            except Exception as exc:
                last_error = exc
                time.sleep(0.6 + attempt * 0.3)
        if matching is None:
            raise RuntimeError(f"PowerPoint did not reopen the staged result: {output}") from last_error
    matching.Windows(1).Activate()
    _maximize_window(int(matching.Windows(1).Hwnd))
    return matching


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
    document.SaveAs2(FileName=str(source), FileFormat=12)
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
        # ppLayoutText gives the fixture real title/body placeholders.  The
        # showcase should prove that the CLI edits PowerPoint's native fields,
        # not merely add small freeform textboxes over a blank slide.
        slide = presentation.Slides.Add(index, 2)  # ppLayoutText
        placeholders = {
            int(shape.PlaceholderFormat.Type): shape
            for shape in slide.Shapes
            if int(shape.Type) == 14
        }
        placeholders[1].TextFrame.TextRange.Text = title_text
        placeholders[2].TextFrame.TextRange.Text = body_text
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
    live: bool = False,
    display_pause: float | None = None,
) -> dict[str, Any]:
    read_operations = WORD_COM_READ_OPERATIONS if app == "word" else SLIDES_COM_READ_OPERATIONS
    if operation in read_operations:
        policy = "read-only"
        write = {"mode": "none"}
    elif live:
        policy = "edit"
        write = {
            "mode": "in-place",
            "snapshot": str((root / "snapshots" / f"{app}-{sequence:03d}.{'docx' if app == 'word' else 'pptx'}").resolve()),
        }
    else:
        policy = "review"
        output = _output_for(app, operation, root, sequence)
        output.parent.mkdir(parents=True, exist_ok=True)
        write = {"mode": "save-as", "path": str(output.resolve())}
    plan_path = root / "plans" / f"{app}-{sequence:03d}-{operation}.json"
    plan_path.parent.mkdir(parents=True, exist_ok=True)
    plan = {
        "schema": "white-collar.plan/v1",
        "app": app,
        "target": {"path": str(target.resolve())},
        "policy": policy,
        "operations": [{"op": operation, "args": args}],
        "write": write,
    }
    if display_pause is not None:
        plan["display"] = {"pause_after_operation": display_pause, "keep_live_as_output": live}
    _write_plan(plan_path, plan)
    value = _run_cli([app, "apply", "--plan", str(plan_path), "--backend", "com"])
    time.sleep(pause)
    return value


def _apply_batch(
    app: str,
    target: Path,
    operations: list[tuple[str, dict[str, Any]]],
    root: Path,
    sequence: int,
    *,
    live: bool = False,
    display_pause: float | None = None,
    keep_live: bool = False,
) -> dict[str, Any]:
    """Run stateful operations in one CLI process (snapshot/diff/undo)."""

    read_operations = WORD_COM_READ_OPERATIONS if app == "word" else SLIDES_COM_READ_OPERATIONS
    is_mutation = any(operation not in read_operations for operation, _args in operations)
    policy = "edit" if live and is_mutation else ("review" if is_mutation else "read-only")
    write = {"mode": "none"}
    if is_mutation:
        if live:
            snapshot = root / "snapshots" / f"{app}-{sequence:03d}.{'docx' if app == 'word' else 'pptx'}"
            snapshot.parent.mkdir(parents=True, exist_ok=True)
            write = {"mode": "in-place", "snapshot": str(snapshot.resolve())}
        else:
            output = _output_for(app, operations[0][0], root, sequence)
            output.parent.mkdir(parents=True, exist_ok=True)
            write = {"mode": "save-as", "path": str(output.resolve())}
    plan_path = root / "plans" / f"{app}-{sequence:03d}-batch.json"
    plan_path.parent.mkdir(parents=True, exist_ok=True)
    plan = {
        "schema": "white-collar.plan/v1",
        "app": app,
        "target": {"path": str(target.resolve())},
        "policy": policy,
        "operations": [{"op": operation, "args": args} for operation, args in operations],
        "write": write,
    }
    if display_pause is not None:
        plan["display"] = {"pause_after_operation": display_pause, "keep_live_as_output": keep_live}
    _write_plan(plan_path, plan)
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


def _legacy_word_mutation_showcase(target: Path, root: Path, assets: dict[str, Path]) -> None:
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


def _word_mutation_showcase(app: Any, target: Path, root: Path, assets: dict[str, Path]) -> None:
    """Run Word in visible, cumulative chapters through the public CLI."""

    current_target = target.resolve()
    sequence = 1

    def stage(operations: list[tuple[str, dict[str, Any]]], *, hold: float = 1.8) -> None:
        nonlocal current_target, sequence
        is_mutation = any(operation not in WORD_COM_READ_OPERATIONS for operation, _args in operations)
        _apply_batch("word", current_target, operations, root, sequence, live=is_mutation)
        sequence += 1
        time.sleep(hold)

    def single(operation: str, args: dict[str, Any], *, hold: float = 1.4) -> None:
        nonlocal sequence
        _apply_operation("word", current_target, operation, args, root, sequence)
        sequence += 1
        time.sleep(hold)

    # The opening shot is intentionally an empty, valid Word document.  The
    # first CLI plan visibly fills it with real content.
    stage(
        [
            (
                "word_live_insert_text",
                {
                    "text": (
                        "Draft target text\rHeading One\rHeading Two\r"
                        "A short paragraph for the white-collar showcase.\r"
                    ),
                    "position": "end",
                },
            )
        ],
        hold=2.5,
    )
    stage(
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
        hold=1.5,
    )
    stage(
        [
            ("word_live_replace_text", {"find_text": "Draft target text", "replace_text": "Reviewed target text", "replace_all": True}),
            ("word_live_insert_paragraphs", {"paragraphs": ["Inserted A", "Inserted B"], "position": "end"}),
            ("word_live_format_text", {"start_paragraph": 1, "end_paragraph": 1, "bold": True, "font_name": "Arial", "font_size": 16}),
            ("word_live_apply_style", {"paragraph_index": 1, "style_name": "Title"}),
            ("word_live_apply_list", {"start_paragraph": 2, "end_paragraph": 3, "list_type": "bullet"}),
            ("word_live_set_paragraph_spacing", {"paragraph_index": 1, "space_after_pt": 6, "line_spacing_rule": "single"}),
        ],
        hold=2.2,
    )
    stage(
        [
            ("word_live_take_snapshot", {}),
            ("word_live_insert_text", {"text": " inserted", "position": "end"}),
            ("word_live_get_diff", {}),
            ("word_live_add_hyperlink", {"paragraph_index": 1, "url": "https://example.com", "display_text": "Example"}),
            ("word_live_list_hyperlinks", {}),
            ("word_live_remove_hyperlink", {"hyperlink_index": 1}),
            ("word_live_list_hyperlinks", {}),
            ("word_live_add_note", {"paragraph_index": 1, "text": "Footnote text", "note_type": "footnote"}),
            ("word_live_list_notes", {}),
            ("word_live_set_content_control", {"target_text": "Reviewed target text", "title": "ClientName", "value": "Example Client", "tag": "client-name"}),
            ("word_live_list_content_controls", {}),
            ("word_live_add_bookmark", {"paragraph_index": 1, "bookmark_name": "CrossRef"}),
            ("word_live_list_cross_reference_items", {}),
            ("word_live_insert_cross_reference", {"ref_type": "Bookmark", "ref_item": "CrossRef", "paragraph_index": 1}),
        ],
        hold=2.2,
    )
    stage(
        [
            ("word_live_add_table", {"rows": 2, "cols": 2, "data": [["A", "B"], ["C", "D"]]}),
            ("word_live_format_table", {"table_index": -1, "autofit": "window", "table_alignment": "center"}),
            ("word_live_modify_table", {"table_index": 1, "operation": "set_cell", "row": 1, "col": 1, "text": "Changed"}),
            ("word_live_insert_image", {"image_path": str(assets["image"]), "position": "end", "width_pt": 72, "height_pt": 40}),
            ("word_live_insert_equation", {"equation": "x = y", "position": "end"}),
        ],
        hold=2.2,
    )
    stage(
        [
            ("word_live_add_comment", {"start": 6, "end": 18, "text": "Review this"}),
            ("word_live_get_comments", {}),
            ("word_live_save", {}),
        ],
        hold=2.0,
    )
    stage(
        [
            ("word_live_add_comment", {"start": 20, "end": 32, "text": "Second review"}),
            ("word_live_get_comments", {}),
            ("word_live_reply_to_comment", {"comment_index": 2, "text": "Noted"}),
            ("word_live_resolve_comment", {"comment_index": 2, "resolve": True}),
            ("word_live_save", {}),
        ],
        hold=2.0,
    )
    stage([("word_live_delete_comment", {"comment_index": 2})], hold=1.5)
    stage(
        [
            ("word_live_replace_text", {"find_text": "Heading One", "replace_text": "Tracked", "replace_all": True, "track_changes": True}),
            ("word_live_list_revisions", {}),
            ("word_live_accept_revisions", {}),
            ("word_live_replace_text", {"find_text": "Heading Two", "replace_text": "Rejected", "replace_all": True, "track_changes": True}),
            ("word_live_reject_revisions", {}),
            ("word_live_toggle_track_changes", {"enable": True}),
        ],
        hold=2.0,
    )
    stage(
        [
            ("word_live_set_protection", {"protection_type": "read_only"}),
            ("word_live_get_protection", {}),
            ("word_live_set_protection", {"protection_type": "none"}),
        ],
        hold=1.8,
    )
    stage(
        [
            ("word_live_set_page_layout", {"orientation": "portrait", "page_width_inches": 8.5, "page_height_inches": 11}),
            ("word_live_add_header_footer", {"header_text": "White-collar showcase", "footer_text": "Live COM"}),
            ("word_live_add_watermark", {"text": "DRAFT"}),
            ("word_live_add_page_numbers", {"position": "footer", "alignment": "center", "prefix": "Page "}),
            ("word_live_add_page_numbers", {"position": "header", "alignment": "right", "prefix": "Page ", "include_total": True, "suffix": " total"}),
        ],
        hold=2.5,
    )
    stage(
        [
            ("word_live_remove_watermark", {"text": "DRAFT"}),
            ("word_live_remove_header_footer", {"position": "both", "section_index": 1}),
            ("word_live_add_section_break", {"break_type": "new_page"}),
            ("word_live_insert_toc", {"position": "start", "lower_heading_level": 2}),
            ("word_live_update_fields", {}),
        ],
        hold=2.2,
    )
    stage(
        [
            ("word_live_setup_heading_numbering", {"h1_paragraphs": [2], "h2_paragraphs": [3]}),
            ("word_live_get_undo_history", {}),
            ("word_live_set_core_properties", {"title": "White-collar showcase", "author": "white-collar"}),
            ("word_live_save", {}),
            ("word_live_undo", {"times": 1}),
        ],
        hold=2.0,
    )
    single("word_live_compare_documents", {"source_path": str(assets["compare"])}, hold=1.2)
    single("word_live_export_pdf", {}, hold=1.2)
    stage(
        [
            ("word_live_merge_document", {"source_path": str(assets["merge"])}),
            ("word_live_delete_text", {"target_text": "Merged source marker"}),
        ],
        hold=2.0,
    )


def _slides_read_showcase(target: Path, root: Path, sequence: int = 1) -> None:
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
        sequence,
    )


def _legacy_slides_mutation_showcase(target: Path, root: Path, assets: dict[str, Path]) -> None:
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
        ("slides_live_distribute", {"slide_index": 4, "shape_names": ["Accent", "Logo", "Body"], "direction": "horizontal"}),
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


def _slides_mutation_showcase(app: Any, presentation: Any, target: Path, root: Path, assets: dict[str, Path]) -> None:
    """Run PowerPoint in visible chapters and focus the changed slide."""

    current_presentation = presentation
    current_target = target.resolve()
    sequence = 1
    read_operations = SLIDES_COM_READ_OPERATIONS

    def stage(
        operations: list[tuple[str, dict[str, Any]]],
        *,
        slide_index: int | None = None,
        after_slide_index: int | None = None,
        hold: float = 1.8,
    ) -> None:
        nonlocal sequence
        if slide_index is not None:
            _show_slide(app, current_presentation, slide_index)
        is_mutation = any(operation not in read_operations for operation, _args in operations)
        _apply_batch("slides", current_target, operations, root, sequence, live=is_mutation)
        if after_slide_index is not None:
            _show_slide(app, current_presentation, after_slide_index)
        elif slide_index is not None:
            _show_slide(app, current_presentation, slide_index)
        sequence += 1
        time.sleep(hold)

    def single(
        operation: str,
        args: dict[str, Any],
        *,
        slide_index: int | None = None,
        after_slide_index: int | None = None,
        hold: float = 1.4,
        live: bool = False,
    ) -> dict[str, Any]:
        nonlocal sequence
        if slide_index is not None:
            _show_slide(app, current_presentation, slide_index)
        value = _apply_operation("slides", current_target, operation, args, root, sequence, live=live)
        if after_slide_index is not None:
            _show_slide(app, current_presentation, after_slide_index)
        elif slide_index is not None:
            _show_slide(app, current_presentation, slide_index)
        sequence += 1
        time.sleep(hold)
        return value

    # Make the first visible change on slide 1, then run reads against the
    # resulting presentation so the recording starts with a real CLI edit.
    stage(
        [
            ("slides_live_insert_text", {"slide_index": 1, "shape_name": "Body", "text": " inserted"}),
            ("slides_live_replace_text", {"find_text": "Draft", "replace_text": "Final", "replace_all": True}),
        ],
        slide_index=1,
        hold=2.2,
    )
    _slides_read_showcase(current_target, root, sequence)
    sequence += 1
    _show_slide(app, current_presentation, 1)
    time.sleep(1.5)

    stage(
        [("slides_live_apply_template", {"source_path": str(assets["template"])})],
        slide_index=1,
        hold=2.0,
    )
    single("slides_live_save_template", {}, slide_index=1, hold=1.0)
    stage(
        [
            ("slides_live_add_slide", {"slide_index": 3, "title": "New Slide"}),
            ("slides_live_set_title", {"slide_index": 3, "title": "Reviewed Slide"}),
            ("slides_live_add_textbox", {"slide_index": 3, "name": "Inserted Text", "text": "Added body", "top": 130}),
            ("slides_live_format_text", {"slide_index": 3, "shape_name": "Inserted Text", "font_name": "Arial", "font_size": 24, "bold": True}),
            ("slides_live_set_layout", {"slide_index": 3, "layout": 1}),
        ],
        slide_index=3,
        hold=2.5,
    )
    stage(
        [
            ("slides_live_add_shape", {"slide_index": 3, "name": "Accent", "shape_type": "rectangle", "fill_color": "204060"}),
            ("slides_live_add_image", {"slide_index": 3, "name": "Logo", "image_path": str(assets["image"]), "top": 240, "width": 120, "height": 75}),
            ("slides_live_set_background", {"slide_index": 3, "color": "F0F4F8"}),
        ],
        slide_index=3,
        hold=2.2,
    )
    stage(
        [
            ("slides_live_duplicate_slide", {"slide_index": 3}),
            ("slides_live_reorder_slide", {"slide_index": 4, "to_index": 2}),
        ],
        slide_index=4,
        hold=2.0,
    )
    stage(
        [
            ("slides_live_set_notes", {"slide_index": 4, "text": "Review notes from the CLI"}),
            ("slides_live_get_notes", {"slide_index": 4}),
            ("slides_live_set_slide_size", {"width_inches": 10, "height_inches": 5.625}),
        ],
        slide_index=4,
        hold=2.0,
    )
    stage(
        [
            ("slides_live_add_shape", {"slide_index": 4, "name": "GroupA", "shape_type": "rectangle", "left": 100, "top": 350}),
            ("slides_live_add_shape", {"slide_index": 4, "name": "GroupB", "shape_type": "rectangle", "left": 300, "top": 350}),
        ],
        slide_index=4,
        hold=2.0,
    )
    group_result = single(
        "slides_live_group",
        {"slide_index": 4, "shape_names": ["GroupA", "GroupB"]},
        slide_index=4,
        hold=1.8,
        live=True,
    )
    group_name = group_result["data"]["operations"][0]["shape"]
    single(
        "slides_live_ungroup",
        {"slide_index": 4, "shape_name": group_name},
        slide_index=4,
        hold=1.8,
        live=True,
    )
    stage(
        [
            ("slides_live_align", {"slide_index": 4, "shape_names": ["Accent", "Logo"], "alignment": "left"}),
            ("slides_live_distribute", {"slide_index": 4, "shape_names": ["Accent", "Logo", "Inserted Text"], "direction": "horizontal"}),
            ("slides_live_z_order", {"slide_index": 4, "shape_name": "Accent", "command": "bring_to_front"}),
            ("slides_live_crop_image", {"slide_index": 4, "shape_name": "Logo", "left": 2, "top": 2, "right": 1, "bottom": 1}),
            ("slides_live_rotate_shape", {"slide_index": 4, "shape_name": "Accent", "degrees": 15}),
        ],
        slide_index=4,
        hold=2.2,
    )
    stage(
        [
            ("slides_live_add_table", {"slide_index": 4, "name": "DataTable", "rows": 2, "columns": 2, "data": [["A", "B"], ["C", "D"]]}),
            ("slides_live_set_table_cell", {"slide_index": 4, "shape_name": "DataTable", "row": 1, "column": 1, "text": "Updated"}),
            ("slides_live_add_chart", {"slide_index": 4, "name": "DataChart", "chart_type": "column", "title": "Results", "data": [["Category", "Value"], ["A", 1], ["B", 2]]}),
            ("slides_live_add_smartart", {"slide_index": 4, "name": "FlowSmartArt", "layout": 1, "nodes": ["Start"]}),
        ],
        slide_index=4,
        hold=2.5,
    )
    stage(
        [
            ("slides_live_add_media", {"slide_index": 4, "name": "Audio", "media_path": str(assets["audio"])}),
            ("slides_live_add_media", {"slide_index": 4, "name": "Video", "media_path": str(assets["video"])}),
            ("slides_live_get_media", {"slide_index": 4}),
            ("slides_live_set_hyperlink", {"slide_index": 4, "shape_name": "Accent", "url": "https://example.com"}),
            ("slides_live_set_alt_text", {"slide_index": 4, "shape_name": "Logo", "text": "Blue logo"}),
            ("slides_live_set_transition", {"slide_index": 4, "effect": "fade", "advance_on_click": True}),
            ("slides_live_add_animation", {"slide_index": 4, "shape_name": "Accent", "effect": "fade"}),
        ],
        slide_index=4,
        hold=2.5,
    )
    stage(
        [("slides_live_add_section", {"name": "Review"}), ("slides_live_get_sections", {})],
        slide_index=4,
        hold=1.7,
    )
    stage(
        [("slides_live_delete_section", {"section_index": 1})],
        slide_index=4,
        hold=1.7,
    )
    stage(
        [("slides_live_set_slide_numbers", {"visible": True})],
        slide_index=4,
        hold=1.5,
    )
    stage(
        [("slides_live_set_slide_visibility", {"slide_index": 4, "visible": False})],
        slide_index=4,
        after_slide_index=1,
        hold=1.8,
    )
    single("slides_live_export_pdf", {}, slide_index=1, hold=1.0)
    single("slides_live_save", {}, slide_index=1, hold=1.8)
    stage(
        [("slides_live_delete_slide", {"slide_index": 2})],
        slide_index=2,
        after_slide_index=1,
        hold=2.0,
    )


def _word_continuous_showcase(app: Any, target: Path, root: Path, assets: dict[str, Path]) -> None:
    """Run the Word matrix in short CLI chapters without closing the live doc."""

    del app
    current_target = target.resolve()
    sequence = 1

    def chapter(operations: list[tuple[str, dict[str, Any]]], *, hold: float = 1.8) -> None:
        nonlocal current_target, sequence
        _apply_batch(
            "word",
            current_target,
            operations,
            root,
            sequence,
            display_pause=0.45,
            keep_live=True,
        )
        if any(operation not in WORD_COM_READ_OPERATIONS for operation, _args in operations):
            current_target = _output_for("word", operations[0][0], root, sequence).resolve()
        sequence += 1
        time.sleep(hold)

    chapter(
        [
            ("word_live_insert_text", {"text": "Draft target text\r", "position": "end"}),
            ("word_live_insert_text", {"text": "Heading One\r", "position": "end"}),
            ("word_live_insert_text", {"text": "Heading Two\r", "position": "end"}),
            ("word_live_insert_text", {"text": "A short paragraph for the white-collar showcase.\r", "position": "end"}),
        ],
        hold=2.2,
    )
    chapter(
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
        hold=1.2,
    )
    chapter(
        [
            ("word_live_replace_text", {"find_text": "Draft target text", "replace_text": "Reviewed target text", "replace_all": True}),
            ("word_live_insert_paragraphs", {"paragraphs": ["Inserted A", "Inserted B"], "position": "end"}),
            ("word_live_format_text", {"start_paragraph": 1, "end_paragraph": 1, "bold": True, "font_name": "Arial", "font_size": 16}),
            ("word_live_apply_style", {"paragraph_index": 1, "style_name": "Title"}),
            ("word_live_apply_list", {"start_paragraph": 2, "end_paragraph": 3, "list_type": "bullet"}),
            ("word_live_set_paragraph_spacing", {"paragraph_index": 1, "space_after_pt": 6, "line_spacing_rule": "single"}),
            ("word_live_take_snapshot", {}),
            ("word_live_insert_text", {"text": " inserted", "position": "end"}),
            ("word_live_get_diff", {}),
        ],
        hold=2.0,
    )
    chapter(
        [
            ("word_live_add_hyperlink", {"paragraph_index": 1, "url": "https://example.com", "display_text": "Example"}),
            ("word_live_list_hyperlinks", {}),
            ("word_live_remove_hyperlink", {"hyperlink_index": 1}),
            ("word_live_list_hyperlinks", {}),
            ("word_live_add_note", {"paragraph_index": 1, "text": "Footnote text", "note_type": "footnote"}),
            ("word_live_list_notes", {}),
            ("word_live_set_content_control", {"target_text": "Reviewed target text", "title": "ClientName", "value": "Example Client", "tag": "client-name"}),
            ("word_live_list_content_controls", {}),
            ("word_live_add_bookmark", {"paragraph_index": 1, "bookmark_name": "CrossRef"}),
            ("word_live_list_cross_reference_items", {}),
            ("word_live_insert_cross_reference", {"ref_type": "Bookmark", "ref_item": "CrossRef", "paragraph_index": 1}),
        ],
        hold=2.0,
    )
    chapter(
        [
            ("word_live_add_table", {"rows": 2, "cols": 2, "data": [["A", "B"], ["C", "D"]]}),
            ("word_live_format_table", {"table_index": -1, "autofit": "window", "table_alignment": "center"}),
            ("word_live_modify_table", {"table_index": 1, "operation": "set_cell", "row": 1, "col": 1, "text": "Changed"}),
            ("word_live_insert_image", {"image_path": str(assets["image"]), "position": "end", "width_pt": 72, "height_pt": 40}),
            ("word_live_insert_equation", {"equation": "x = y", "position": "end"}),
        ],
        hold=2.0,
    )
    chapter(
        [("word_live_add_comment", {"start": 6, "end": 18, "text": "Review this"}), ("word_live_get_comments", {}), ("word_live_save", {})],
        hold=1.8,
    )
    chapter(
        [("word_live_add_comment", {"start": 20, "end": 32, "text": "Second review"}), ("word_live_get_comments", {}), ("word_live_save", {})],
        hold=1.8,
    )
    chapter(
        [("word_live_reply_to_comment", {"comment_index": 2, "text": "Noted"}), ("word_live_resolve_comment", {"comment_index": 2, "resolve": True}), ("word_live_save", {})],
        hold=1.8,
    )
    chapter(
        [
            ("word_live_delete_comment", {"comment_index": 2}),
            ("word_live_replace_text", {"find_text": "Heading One", "replace_text": "Tracked", "replace_all": True, "track_changes": True}),
            ("word_live_list_revisions", {}),
            ("word_live_accept_revisions", {}),
            ("word_live_replace_text", {"find_text": "Heading Two", "replace_text": "Rejected", "replace_all": True, "track_changes": True}),
            ("word_live_reject_revisions", {}),
            ("word_live_toggle_track_changes", {"enable": True}),
            ("word_live_set_protection", {"protection_type": "read_only"}),
            ("word_live_get_protection", {}),
            ("word_live_set_protection", {"protection_type": "none"}),
        ],
        hold=2.0,
    )
    chapter(
        [
            ("word_live_set_page_layout", {"orientation": "portrait", "page_width_inches": 8.5, "page_height_inches": 11}),
            ("word_live_add_header_footer", {"header_text": "White-collar showcase", "footer_text": "Live COM"}),
            ("word_live_add_watermark", {"text": "DRAFT"}),
            ("word_live_add_page_numbers", {"position": "footer", "alignment": "center", "prefix": "Page "}),
            ("word_live_add_page_numbers", {"position": "header", "alignment": "right", "prefix": "Page ", "include_total": True, "suffix": " total"}),
        ],
        hold=2.2,
    )
    chapter(
        [
            ("word_live_remove_watermark", {"text": "DRAFT"}),
            ("word_live_remove_header_footer", {"position": "both", "section_index": 1}),
            ("word_live_add_section_break", {"break_type": "new_page"}),
            ("word_live_insert_toc", {"position": "start", "lower_heading_level": 2}),
            ("word_live_update_fields", {}),
            ("word_live_setup_heading_numbering", {"h1_paragraphs": [2], "h2_paragraphs": [3]}),
        ],
        hold=2.0,
    )
    chapter(
        [
            ("word_live_get_undo_history", {}),
            ("word_live_set_core_properties", {"title": "White-collar showcase", "author": "white-collar"}),
            ("word_live_save", {}),
            ("word_live_undo", {"times": 1}),
            ("word_live_merge_document", {"source_path": str(assets["merge"])}),
            ("word_live_delete_text", {"target_text": "Merged source marker"}),
        ],
        hold=2.0,
    )
    _apply_operation(
        "word",
        current_target,
        "word_live_compare_documents",
        {"source_path": str(assets["compare"])},
        root,
        sequence,
        display_pause=0.25,
    )
    _apply_operation("word", current_target, "word_live_export_pdf", {}, root, sequence + 1, display_pause=0.25)


def _slides_continuous_showcase(app: Any, presentation: Any, target: Path, root: Path, assets: dict[str, Path]) -> None:
    """Run PowerPoint in three continuous CLI plans with slide-aware pauses."""

    template_result = _apply_operation(
        "slides",
        target,
        "slides_live_save_template",
        {},
        root,
        1,
    )
    del template_result
    main_operations = [
        ("slides_live_insert_text", {"slide_index": 1, "shape_name": "Body", "text": " inserted"}),
        ("slides_live_replace_text", {"find_text": "Draft", "replace_text": "Final", "replace_all": True}),
        ("slides_live_list_open", {}),
        ("slides_live_get_info", {}),
        ("slides_live_get_text", {}),
        ("slides_live_get_slide_text", {"slide_index": 1}),
        ("slides_live_find_text", {"search_text": "Final"}),
        ("slides_live_get_masters", {}),
        ("slides_live_get_layouts", {}),
        ("slides_live_get_placeholders", {"master": 1}),
        ("slides_live_get_notes", {"slide_index": 1}),
        ("slides_live_get_sections", {}),
        ("slides_live_get_media", {}),
        ("slides_live_apply_template", {"source_path": str(assets["template"])}),
        ("slides_live_add_slide", {"slide_index": 3, "title": "New Slide"}),
        ("slides_live_set_layout", {"slide_index": 3, "layout": 2}),
        ("slides_live_set_title", {"slide_index": 3, "title": "Reviewed Slide"}),
        ("slides_live_insert_text", {"slide_index": 3, "shape_name": "Body", "text": "Added body\r"}),
        ("slides_live_format_text", {"slide_index": 3, "shape_name": "Body", "font_name": "Arial", "font_size": 24, "bold": True}),
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
    _apply_batch("slides", target, main_operations, root, 2, display_pause=0.55)
    main_output = _output_for("slides", main_operations[0][0], root, 2).resolve()
    # PowerPoint's SaveCopyAs leaves the edited presentation live.  Continue
    # against that same window instead of reopening the review copy.
    _show_slide(app, presentation, 4)
    group_result = _apply_operation(
        "slides",
        target,
        "slides_live_group",
        {"slide_index": 4, "shape_names": ["GroupA", "GroupB"]},
        root,
        3,
        display_pause=0.8,
    )
    group_name = group_result["data"]["operations"][0]["shape"]
    grouped_output = target.resolve()
    _show_slide(app, presentation, 4)
    remaining_operations = [
        ("slides_live_ungroup", {"slide_index": 4, "shape_name": group_name}),
        ("slides_live_align", {"slide_index": 4, "shape_names": ["Accent", "Logo"], "alignment": "left"}),
        ("slides_live_distribute", {"slide_index": 4, "shape_names": ["Accent", "Logo", "Body"], "direction": "horizontal"}),
        ("slides_live_z_order", {"slide_index": 4, "shape_name": "Accent", "command": "bring_to_front"}),
        ("slides_live_crop_image", {"slide_index": 4, "shape_name": "Logo", "left": 2, "top": 2, "right": 1, "bottom": 1}),
        ("slides_live_rotate_shape", {"slide_index": 4, "shape_name": "Accent", "degrees": 15}),
        ("slides_live_add_section", {"name": "Review"}),
        ("slides_live_get_sections", {}),
        ("slides_live_delete_section", {"section_index": 1}),
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
        ("slides_live_set_slide_numbers", {"visible": True}),
        ("slides_live_set_slide_visibility", {"slide_index": 4, "visible": False}),
        ("slides_live_save", {}),
        ("slides_live_delete_slide", {"slide_index": 2}),
    ]
    _apply_batch("slides", grouped_output, remaining_operations, root, 4, display_pause=0.65)
    _show_slide(app, presentation, 1)
    _apply_operation("slides", grouped_output, "slides_live_export_pdf", {}, root, 5, display_pause=0.25)
    time.sleep(2.5)


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
            word_clip = root / "word-showcase.mp4"
            _capture_segment(
                ffmpeg,
                word_hwnd,
                word_clip,
                lambda: _word_continuous_showcase(word_app, word_source, root, word_assets),
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
            slides_clip = root / "slides-showcase.mp4"
            _capture_segment(
                ffmpeg,
                slides_hwnd,
                slides_clip,
                lambda: _slides_continuous_showcase(slides_app, slides_presentation, slides_source, root, slides_assets),
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
