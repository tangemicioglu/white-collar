from __future__ import annotations

import os
import uuid
import zipfile
from pathlib import Path

import pytest

from whitecollar.adapters.slides import PowerPointComAdapter
from whitecollar.models import Plan
from whitecollar.slides_ops import SLIDES_COM_MUTATING_OPERATIONS, SLIDES_COM_OPERATIONS


pytestmark = pytest.mark.real_powerpoint
_REAL_POWERPOINT_BEFORE: set[int] = set()


def _plan(target: Path, operation: str, args: dict, root: Path, sequence: int) -> Plan:
    if operation in {"slides_live_list_open", "slides_live_get_info", "slides_live_get_text", "slides_live_get_slide_text", "slides_live_find_text", "slides_live_get_masters", "slides_live_get_layouts", "slides_live_get_placeholders", "slides_live_get_notes", "slides_live_get_sections", "slides_live_get_media"}:
        policy = "read-only"
        write = {"mode": "none"}
    elif operation == "slides_live_save_template":
        policy = "edit"
        write = {"mode": "save-as", "path": str(root / f"template-{sequence:03d}.potx")}
    elif operation == "slides_live_export_pdf":
        policy = "edit"
        write = {"mode": "save-as", "path": str(root / f"export-{sequence:03d}.pdf")}
    elif operation == "slides_screen_capture":
        policy = "review"
        write = {"mode": "save-as", "path": str(root / f"screen-source-{sequence:03d}.pptx")}
    else:
        policy = "edit"
        write = {"mode": "in-place", "snapshot": str(root / f"snapshot-{sequence:02d}-{operation}.pptx")}
    return Plan.from_dict(
        {
            "schema": "white-collar.plan/v1",
            "app": "slides",
            "target": {"path": str(target)},
            "policy": policy,
            "operations": [{"op": operation, "args": args}],
            "write": write,
        }
    )


def _add_textbox(slide, name: str, text: str, *, top: float) -> object:
    shape = slide.Shapes.AddTextbox(1, 60, top, 600, 70)
    shape.Name = name
    shape.TextFrame.TextRange.Text = text
    return shape


def _new_presentation(app, path: Path):
    presentation = app.Presentations.Add(True)
    first = presentation.Slides.Add(1, 12)
    _add_textbox(first, "Title", "Draft Review", top=40)
    _add_textbox(first, "Body", "Draft status\rAgenda", top=130)
    second = presentation.Slides.Add(2, 12)
    _add_textbox(second, "Title", "Second Slide", top=40)
    _add_textbox(second, "Body", "Planning notes", top=130)
    presentation.SaveAs(FileName=str(path), FileFormat=24)
    return presentation


def _current_presentation(app, target: Path):
    wanted = target.resolve()
    for presentation in list(app.Presentations):
        if Path(str(presentation.FullName)).resolve() == wanted:
            return presentation
    raise AssertionError(f"target is not open in PowerPoint: {target}")


def _shape(presentation, slide_index: int, name: str):
    slide = presentation.Slides(slide_index)
    for shape in list(slide.Shapes):
        if str(shape.Name) == name:
            return shape
    raise AssertionError(f"shape {name!r} not found on slide {slide_index}")


def _slide_text(presentation, slide_index: int) -> str:
    values = []
    for shape in list(presentation.Slides(slide_index).Shapes):
        try:
            if bool(shape.HasTextFrame):
                values.append(str(shape.TextFrame.TextRange.Text))
        except Exception:
            pass
    return "\n".join(values)


def _assert_valid_copy(app, path: Path) -> None:
    assert path.is_file(), path
    assert zipfile.is_zipfile(path), path
    opened = app.Presentations.Open(FileName=str(path), ReadOnly=True, Untitled=False, WithWindow=False)
    try:
        assert int(opened.Slides.Count) >= 1
    finally:
        opened.Close()


def _assert_screenshot(path: Path) -> None:
    from PIL import Image, ImageStat

    assert path.is_file(), path
    with Image.open(path) as image:
        image.load()
        assert image.width > 100 and image.height > 100
        assert max(ImageStat.Stat(image.convert("RGB")).stddev) > 0


def _assert_pdf(path: Path) -> None:
    assert path.is_file(), path
    assert path.read_bytes()[:5] == b"%PDF-"


def _assert_rendered_slide(path: Path) -> None:
    from PIL import Image, ImageStat

    assert path.is_file(), path
    with Image.open(path) as image:
        image.load()
        assert image.width >= 640 and image.height >= 360
        assert max(ImageStat.Stat(image.convert("RGB")).stddev) > 0


def _artifact_root(tmp_path: Path) -> Path:
    configured = os.environ.get("WHITE_COLLAR_REAL_POWERPOINT_ARTIFACT_DIR")
    if not configured:
        return tmp_path / "real-powerpoint-artifacts"
    root = Path(configured).resolve() / f"run-{uuid.uuid4().hex[:10]}"
    root.mkdir(parents=True, exist_ok=False)
    return root


def _new_template(app, path: Path) -> None:
    template = app.Presentations.Add(True)
    try:
        template.Slides.Add(1, 12)
        template.SaveAs(FileName=str(path), FileFormat=26)
    finally:
        template.Close()


def _assert_operation_behavior(app, target: Path, operation: str, args: dict, value: dict) -> None:
    result = value["operations"][0]
    presentation = _current_presentation(app, target)
    if operation == "slides_live_list_open":
        assert result["presentations"]
    elif operation == "slides_live_get_info":
        assert result["path"] == str(target)
        assert result["slides"] >= 2
        assert result["slide_width"] > 0 and result["slide_height"] > 0
    elif operation == "slides_live_get_text":
        assert any("Draft Review" in shape["text"] for slide in result["slides"] for shape in slide["shapes"])
    elif operation == "slides_live_get_slide_text":
        assert "Draft status" in " ".join(shape["text"] for shape in result["shapes"])
    elif operation == "slides_live_find_text":
        assert result["matches"] and result["matches"][0]["slide_index"] == 1
    elif operation == "slides_live_insert_text":
        assert "inserted" in _slide_text(presentation, int(args.get("slide_index", 1)))
    elif operation in {"slides_live_replace_text", "replace_text"}:
        assert "Final" in _slide_text(presentation, 1)
        assert result["replacements"] >= 1
    elif operation == "slides_live_add_slide":
        assert presentation.Slides.Count >= 3
        assert "New Slide" in _slide_text(presentation, 3)
    elif operation == "slides_live_delete_slide":
        assert presentation.Slides.Count == 3
    elif operation == "slides_live_set_title":
        assert "Reviewed Slide" in _slide_text(presentation, int(args.get("slide_index", 3)))
    elif operation == "slides_live_add_textbox":
        assert "Added body" in _slide_text(presentation, int(args.get("slide_index", 3)))
    elif operation == "slides_live_format_text":
        shape = _shape(presentation, int(args.get("slide_index", 3)), str(args["shape_name"]))
        font = shape.TextFrame.TextRange.Font
        assert str(font.Name) == "Arial"
        assert abs(float(font.Size) - 24) < 0.1
        assert bool(font.Bold) is True
    elif operation == "slides_live_add_shape":
        shape = _shape(presentation, int(args.get("slide_index", 3)), str(args.get("name", "Accent")))
        assert int(shape.Type) == 1
    elif operation == "slides_live_add_image":
        shape = _shape(presentation, int(args.get("slide_index", 3)), "Logo")
        assert int(shape.Type) == 13
    elif operation == "slides_live_set_background":
        slide = presentation.Slides(int(args.get("slide_index", 3)))
        assert bool(slide.FollowMasterBackground) is False
    elif operation == "slides_live_duplicate_slide":
        assert presentation.Slides.Count == 4
    elif operation == "slides_live_reorder_slide":
        assert int(presentation.Slides(int(args["to_index"])).SlideID) == int(result["slide_id"])
    elif operation == "slides_live_set_notes":
        notes_text = " ".join(
            str(shape.TextFrame.TextRange.Text)
            for shape in presentation.Slides(int(args.get("slide_index", 3))).NotesPage.Shapes
            if bool(getattr(shape, "HasTextFrame", False))
        )
        assert "Review notes" in notes_text
    elif operation == "slides_live_set_slide_size":
        assert abs(float(presentation.PageSetup.SlideWidth) - 720) < 0.5
        assert abs(float(presentation.PageSetup.SlideHeight) - 405) < 0.5
    elif operation == "slides_live_get_masters":
        assert result["count"] >= 1
    elif operation == "slides_live_get_layouts":
        assert result["count"] >= 1
    elif operation == "slides_live_get_placeholders":
        assert "placeholders" in result
    elif operation == "slides_live_get_notes":
        assert result["notes"]
    elif operation == "slides_live_get_sections":
        assert result["count"] >= 0
    elif operation == "slides_live_get_media":
        assert "media" in result
    elif operation == "slides_live_apply_template":
        assert result["applied"] is True
    elif operation == "slides_live_save_template":
        assert Path(result["path"]).is_file()
        assert Path(result["path"]).read_bytes()[:2] == b"PK"
    elif operation == "slides_live_set_layout":
        assert str(presentation.Slides(int(args.get("slide_index", 1))).CustomLayout.Name) == str(result["layout"])
    elif operation == "slides_live_group":
        assert int(_shape(presentation, int(args.get("slide_index", 1)), result["shape"]).Type) == 6
    elif operation == "slides_live_ungroup":
        assert result["ungrouped"] is True
    elif operation in {"slides_live_align", "slides_live_distribute", "slides_live_z_order", "slides_live_crop_image", "slides_live_rotate_shape"}:
        assert result["op"] == operation
    elif operation == "slides_live_add_section":
        assert result["count"] >= 1
    elif operation == "slides_live_delete_section":
        assert result["count"] >= 0
    elif operation == "slides_live_set_slide_visibility":
        assert bool(presentation.Slides(int(args.get("slide_index", 1))).SlideShowTransition.Hidden) is True
    elif operation == "slides_live_set_slide_numbers":
        assert result["visible"] is True
    elif operation == "slides_live_add_table":
        assert _shape(presentation, int(args.get("slide_index", 1)), "DataTable").HasTable
    elif operation == "slides_live_set_table_cell":
        table = _shape(presentation, int(args.get("slide_index", 1)), "DataTable").Table
        assert "Updated" in str(table.Cell(1, 1).Shape.TextFrame.TextRange.Text)
    elif operation == "slides_live_add_chart":
        shape = _shape(presentation, int(args.get("slide_index", 1)), "DataChart")
        assert shape.HasChart
        chart = shape.Chart
        if args.get("data"):
            assert int(chart.SeriesCollection().Count) >= 1
    elif operation == "slides_live_add_smartart":
        smartart = _shape(presentation, int(args.get("slide_index", 1)), "FlowSmartArt")
        assert smartart.Type == 24
        if args.get("nodes"):
            assert str(smartart.SmartArt.AllNodes(1).TextFrame2.TextRange.Text).strip() == args["nodes"][0]
    elif operation == "slides_live_add_media":
        assert _shape(presentation, int(args.get("slide_index", 1)), "Audio").Type == 16
    elif operation == "slides_live_set_hyperlink":
        assert str(_shape(presentation, int(args.get("slide_index", 1)), "Accent").ActionSettings(1).Hyperlink.Address).rstrip("/") == str(args["url"]).rstrip("/")
    elif operation == "slides_live_set_alt_text":
        assert str(_shape(presentation, int(args.get("slide_index", 1)), "Logo").AlternativeText) == str(args["text"])
    elif operation == "slides_live_set_transition":
        assert result["op"] == operation
        assert int(presentation.Slides(int(args.get("slide_index", 1))).SlideShowTransition.EntryEffect) != 0
    elif operation == "slides_live_add_animation":
        assert result["op"] == operation
        assert int(presentation.Slides(int(args.get("slide_index", 1))).TimeLine.MainSequence.Count) >= 1
    elif operation == "slides_live_export_pdf":
        _assert_pdf(Path(result["path"]))
    elif operation == "slides_live_save":
        assert bool(presentation.Saved) is True
    elif operation == "slides_screen_capture":
        _assert_screenshot(Path(args["output_path"]))


def _powerpoint_process_ids() -> set[int]:
    import win32gui
    import win32process

    process_ids: set[int] = set()

    def collect(hwnd: int, _extra: object) -> None:
        if win32gui.GetClassName(hwnd) == "PPTFrameClass":
            process_ids.add(int(win32process.GetWindowThreadProcessId(hwnd)[1]))

    win32gui.EnumWindows(collect, None)
    return process_ids


@pytest.fixture(scope="module")
def real_powerpoint():
    from win32com.client import DispatchEx

    global _REAL_POWERPOINT_BEFORE
    _REAL_POWERPOINT_BEFORE = _powerpoint_process_ids()
    app = DispatchEx("PowerPoint.Application")
    app.DisplayAlerts = 0
    try:
        yield app
    finally:
        try:
            presentations = list(app.Presentations)
        except Exception:
            presentations = []
        for presentation in presentations:
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass


def test_every_registered_slides_operation_against_real_powerpoint(real_powerpoint, tmp_path):
    document_path = tmp_path / "real-powerpoint.pptx"
    presentation = _new_presentation(real_powerpoint, document_path)
    owned = _powerpoint_process_ids() - _REAL_POWERPOINT_BEFORE
    if owned:
        os.environ["WHITE_COLLAR_REAL_POWERPOINT_PID"] = str(sorted(owned)[-1])
    adapter = PowerPointComAdapter(app_factory=lambda: real_powerpoint)
    target = Path(str(presentation.FullName))
    artifact_root = _artifact_root(tmp_path)
    screenshot_root = artifact_root / "screenshots"
    rendered_root = artifact_root / "rendered-slides"
    screenshot_root.mkdir(parents=True, exist_ok=True)
    rendered_root.mkdir(parents=True, exist_ok=True)

    image_path = artifact_root / "pixel.png"
    from PIL import Image
    import wave

    Image.new("RGB", (24, 24), (32, 128, 224)).save(image_path)
    audio_path = artifact_root / "tone.wav"
    with wave.open(str(audio_path), "wb") as audio:
        audio.setnchannels(1)
        audio.setsampwidth(2)
        audio.setframerate(8000)
        audio.writeframes(b"\x00\x00" * 800)
    template_path = artifact_root / "source.potx"
    _new_template(real_powerpoint, template_path)
    cases = [
        ("slides_live_list_open", {}),
        ("slides_live_get_info", {}),
        ("slides_live_get_text", {}),
        ("slides_live_get_slide_text", {"slide_index": 1}),
        ("slides_live_find_text", {"search_text": "Draft"}),
        ("slides_live_get_masters", {}),
        ("slides_live_get_layouts", {}),
        ("slides_live_get_placeholders", {"slide_index": 1}),
        ("slides_live_get_notes", {"slide_index": 1}),
        ("slides_live_get_sections", {}),
        ("slides_live_get_media", {}),
        ("slides_live_insert_text", {"slide_index": 1, "shape_name": "Body", "text": " inserted"}),
        ("slides_live_replace_text", {"find_text": "Draft", "replace_text": "Final", "replace_all": True}),
        ("slides_live_apply_template", {"source_path": str(template_path)}),
        ("slides_live_save_template", {}),
        ("slides_live_add_slide", {"slide_index": 3, "title": "New Slide"}),
        ("slides_live_set_title", {"slide_index": 3, "title": "Reviewed Slide"}),
        ("slides_live_add_textbox", {"slide_index": 3, "name": "Inserted Text", "text": "Added body", "top": 130}),
        ("slides_live_format_text", {"slide_index": 3, "shape_name": "Inserted Text", "font_name": "Arial", "font_size": 24, "bold": True}),
        ("slides_live_set_layout", {"slide_index": 3, "layout": 1}),
        ("slides_live_add_shape", {"slide_index": 3, "name": "Accent", "shape_type": "rectangle", "fill_color": "204060"}),
        ("slides_live_add_image", {"slide_index": 3, "name": "Logo", "image_path": str(image_path), "top": 240, "width": 48, "height": 48}),
        ("slides_live_set_background", {"slide_index": 3, "color": "F0F4F8"}),
        ("slides_live_duplicate_slide", {"slide_index": 3}),
        ("slides_live_reorder_slide", {"slide_index": 4, "to_index": 2}),
        ("slides_live_set_notes", {"slide_index": 4, "text": "Review notes"}),
        ("slides_live_set_slide_size", {"width_inches": 10, "height_inches": 5.625}),
        ("slides_live_add_shape", {"slide_index": 4, "name": "GroupA", "shape_type": "rectangle", "left": 100, "top": 350}),
        ("slides_live_add_shape", {"slide_index": 4, "name": "GroupB", "shape_type": "rectangle", "left": 300, "top": 350}),
        ("slides_live_group", {"slide_index": 4, "shape_names": ["GroupA", "GroupB"]}),
        ("slides_live_ungroup", {"slide_index": 4, "shape_name": "__group_from_previous__"}),
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
        ("slides_live_add_media", {"slide_index": 4, "name": "Audio", "media_path": str(audio_path)}),
        ("slides_live_set_hyperlink", {"slide_index": 4, "shape_name": "Accent", "url": "https://example.com"}),
        ("slides_live_set_alt_text", {"slide_index": 4, "shape_name": "Logo", "text": "Blue logo"}),
        ("slides_live_set_transition", {"slide_index": 4, "effect": "fade", "advance_on_click": True}),
        ("slides_live_add_animation", {"slide_index": 4, "shape_name": "Accent", "effect": "fade"}),
        ("slides_live_export_pdf", {}),
        ("slides_live_save", {}),
        ("slides_screen_capture", {"slide_index": 1, "output_path": str(screenshot_root / "final-powerpoint-window.png")}),
        ("slides_live_delete_slide", {"slide_index": 2}),
    ]

    executed = set()
    last_group_name = None
    for sequence, (operation, args) in enumerate(cases, start=1):
        if operation == "slides_live_ungroup" and args.get("shape_name") == "__group_from_previous__":
            args = dict(args)
            args["shape_name"] = last_group_name
        try:
            value = adapter.apply(_plan(target, operation, args, artifact_root, sequence), dry_run=False)
        except Exception as exc:
            pytest.fail(f"real PowerPoint operation {operation} failed: {type(exc).__name__}: {exc}")
        assert value["backend"] == "powerpoint-com", operation
        assert value["operations"], operation
        _assert_operation_behavior(real_powerpoint, target, operation, args, value)
        executed.add(operation)
        if operation == "slides_live_group":
            last_group_name = value["operations"][0]["shape"]

        if operation == "slides_screen_capture":
            _assert_valid_copy(real_powerpoint, artifact_root / f"screen-source-{sequence:03d}.pptx")
        elif operation == "slides_live_save_template":
            assert Path(value["operations"][0]["path"]).is_file()
        elif operation == "slides_live_export_pdf":
            _assert_pdf(Path(value["operations"][0]["path"]))
        elif operation in SLIDES_COM_MUTATING_OPERATIONS:
            _assert_valid_copy(real_powerpoint, artifact_root / f"snapshot-{sequence:02d}-{operation}.pptx")
            screenshot = screenshot_root / f"after-{sequence:02d}-{operation}.png"
            capture = adapter.apply(
                _plan(target, "slides_screen_capture", {"slide_index": min(int(args.get("slide_index", 1)), int(_current_presentation(real_powerpoint, target).Slides.Count)), "output_path": str(screenshot)}, artifact_root, sequence * 1000),
                dry_run=False,
            )
            assert capture["operations"][0]["op"] == "slides_screen_capture"
            _assert_screenshot(screenshot)
            _assert_valid_copy(real_powerpoint, artifact_root / f"screen-source-{sequence * 1000:03d}.pptx")

        current_presentation = _current_presentation(real_powerpoint, target)
        render_index = min(int(args.get("slide_index", 1)), int(current_presentation.Slides.Count))
        rendered = rendered_root / f"after-{sequence:02d}-{operation}.png"
        current_presentation.Slides(render_index).Export(str(rendered), "PNG", 1280, 720)
        _assert_rendered_slide(rendered)

    # Creation has a separate disposable-target test because it does not act
    # on the shared open presentation used by the operation matrix below.
    assert SLIDES_COM_OPERATIONS - executed == {"slides_live_create_presentation"}
    expected_screenshots = 1 + sum(
        operation in SLIDES_COM_MUTATING_OPERATIONS
        and operation not in {"slides_live_save_template", "slides_live_export_pdf"}
        for operation, _ in cases
        if operation != "slides_screen_capture"
    )
    assert len(list(screenshot_root.glob("*.png"))) == expected_screenshots
    presentation = None
    adapter = None


def test_create_powerpoint_presentation_from_scratch_against_real_powerpoint(real_powerpoint, tmp_path):
    target = tmp_path / "created-from-scratch.pptx"
    plan = Plan.from_dict(
        {
            "schema": "white-collar.plan/v1",
            "app": "slides",
            "target": {"path": str(target)},
            "policy": "review",
            "operations": [{"op": "slides_live_create_presentation"}],
            "write": {"mode": "create"},
        }
    )
    adapter = PowerPointComAdapter(app_factory=lambda: real_powerpoint)

    preview = adapter.apply(plan, dry_run=True)
    assert preview["written"] is False
    assert not target.exists()

    value = adapter.apply(plan, dry_run=False)
    assert value["backend"] == "powerpoint-com"
    assert value["operations"][0]["op"] == "slides_live_create_presentation"
    assert value["operations"][0]["created"] is True
    assert value["operations"][0]["slides"] == 1
    _assert_valid_copy(real_powerpoint, target)
