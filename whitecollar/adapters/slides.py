from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any, Callable, Iterator, Protocol

from ..errors import BackendUnavailableError, ValidationError
from ..models import Plan
from ..office_capture import capture_window
from ..slides_ops import SLIDES_COM_MUTATING_OPERATIONS, SLIDES_COM_OPERATIONS


class SlidesAdapter(Protocol):
    def inspect(self, target: Path, *, render_dir: Path | None = None) -> dict[str, Any]: ...

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]: ...


class UnavailableSlidesAdapter:
    def inspect(self, target: Path, *, render_dir: Path | None = None) -> dict[str, Any]:
        raise BackendUnavailableError("slides")

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
        raise BackendUnavailableError("slides")


class PowerPointComAdapter:
    """Finite semantic PowerPoint adapter; plans cannot name arbitrary COM calls."""

    def __init__(
        self,
        *,
        app_factory: Callable[[], Any] | None = None,
        screenshotter: Callable[[int, Path], None] | None = None,
    ) -> None:
        self._app_factory = app_factory or _default_powerpoint_app
        self._screenshotter = screenshotter or capture_window

    def _get_app(self) -> Any:
        return self._app_factory()

    def inspect(self, target: Path, *, render_dir: Path | None = None) -> dict[str, Any]:
        app = self._get_app()
        presentation, opened_here = _find_or_open_for_inspect(app, target)
        try:
            data = self._get_info(presentation) | {"backend": "powerpoint-com"}
            if render_dir is not None:
                data["renders"] = _render_slides(presentation, render_dir)
            return data
        finally:
            if opened_here:
                presentation.Close()

    def apply(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
        if dry_run and all(
            operation["op"] in SLIDES_COM_MUTATING_OPERATIONS or operation["op"] == "replace_text"
            for operation in plan.operations
        ):
            return {
                "backend": "powerpoint-com",
                "written": False,
                "operations": [
                    {"op": operation["op"], "dry_run": True, "args": _operation_args(operation)}
                    for operation in plan.operations
                ],
            }
        app = self._get_app()
        presentation = _find_presentation(app, plan.target.path)
        if not dry_run and plan.write.mode != "none":
            self._prepare_write(presentation, plan)
        operations = []
        for operation in plan.operations:
            name = operation["op"]
            if name not in SLIDES_COM_OPERATIONS and name not in {"replace_text"}:
                raise ValidationError(f"unsupported PowerPoint COM operation: {name}")
            if dry_run and name in SLIDES_COM_MUTATING_OPERATIONS:
                operations.append({"op": name, "dry_run": True, "args": _operation_args(operation)})
                continue
            dispatch_name = "slides_live_replace_text" if name == "replace_text" else name
            method = getattr(self, f"_{dispatch_name}", None)
            if method is None:
                raise ValidationError(f"PowerPoint operation is registered but not implemented: {name}")
            operations.append(method(app, presentation, _operation_args(operation)))
        if not dry_run and plan.write.mode != "none":
            self._commit_write(app, presentation, plan)
        return {"backend": "powerpoint-com", "written": not dry_run, "operations": operations}

    def _prepare_write(self, presentation: Any, plan: Plan) -> None:
        if plan.write.mode != "in-place":
            return
        snapshot = Path(plan.write.snapshot or "")
        if snapshot.exists():
            raise ValidationError("snapshot already exists", details={"snapshot": str(snapshot)})
        snapshot.parent.mkdir(parents=True, exist_ok=True)
        try:
            presentation.SaveCopyAs(str(snapshot))
        except Exception:
            source = Path(str(_safe_value(presentation, "FullName", "")))
            if not source.is_file():
                raise ValidationError(
                    "PowerPoint could not create a snapshot",
                    details={"source": str(source), "snapshot": str(snapshot)},
                )
            if _safe_value(presentation, "Saved", True) is False:
                presentation.Save()
            shutil.copy2(source, snapshot)

    def _commit_write(self, app: Any, presentation: Any, plan: Plan) -> None:
        if plan.write.mode == "in-place":
            presentation.Save()
            return
        output = Path(plan.write.path or "")
        if output.exists():
            raise ValidationError("save-as path already exists", details={"path": str(output)})
        output.parent.mkdir(parents=True, exist_ok=True)
        try:
            presentation.SaveCopyAs(str(output))
        except Exception:
            source = Path(str(_safe_value(presentation, "FullName", "")))
            presentation.SaveAs(str(output), 24)
            if source.resolve() != output.resolve():
                presentation.Close()
                app.Presentations.Open(FileName=str(source), ReadOnly=False, Untitled=False, WithWindow=True)

    def _slides_live_list_open(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        documents = []
        for item in _iter_collection(app.Presentations):
            documents.append(
                {
                    "name": str(_safe_value(item, "Name", "")),
                    "path": str(_safe_value(item, "FullName", "")),
                    "slides": int(_safe_value(item.Slides, "Count", 0)),
                }
            )
        return {"op": "slides_live_list_open", "presentations": documents}

    def _slides_live_get_info(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        return {"op": "slides_live_get_info", **self._get_info(presentation)}

    def _slides_live_get_text(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slides = []
        for index, slide in enumerate(_iter_collection(presentation.Slides), start=1):
            slides.append({"slide_index": index, "shapes": _slide_text(slide)})
        return {"op": "slides_live_get_text", "slides": slides}

    def _slides_live_get_slide_text(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        index = int(args["slide_index"])
        slide = _slide(presentation, index)
        return {"op": "slides_live_get_slide_text", "slide_index": index, "shapes": _slide_text(slide)}

    def _slides_live_find_text(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        search = str(args["search_text"])
        matches = []
        for slide_index, slide in enumerate(_iter_collection(presentation.Slides), start=1):
            for shape in _iter_collection(slide.Shapes):
                text = _shape_text(shape)
                if search.lower() in text.lower():
                    matches.append({"slide_index": slide_index, "shape": str(_safe_value(shape, "Name", "")), "text": text})
        return {"op": "slides_live_find_text", "find": search, "matches": matches}

    def _slides_live_insert_text(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args, required=False)
        if shape is None:
            shape = _add_textbox(slide, args.get("text", ""), args)
        else:
            current = _shape_text(shape)
            value = str(args["text"])
            shape.TextFrame.TextRange.Text = value + current if args.get("position") == "start" else current + value
        return {"op": "slides_live_insert_text", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_replace_text(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        find_text = str(args["find_text"])
        replacement = str(args["replace_text"])
        replace_all = bool(args.get("replace_all", True))
        replacements = 0
        for slide in _iter_collection(presentation.Slides):
            for shape in _iter_collection(slide.Shapes):
                text = _shape_text(shape)
                if find_text.lower() not in text.lower():
                    continue
                if replace_all:
                    replacements += text.count(find_text)
                    shape.TextFrame.TextRange.Text = text.replace(find_text, replacement)
                elif replacements == 0:
                    shape.TextFrame.TextRange.Text = text.replace(find_text, replacement, 1)
                    replacements = 1
        return {"op": "slides_live_replace_text", "find": find_text, "replace": replacement, "replacements": replacements}

    def _slides_live_add_slide(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        index = int(args.get("slide_index", int(presentation.Slides.Count) + 1))
        slide = presentation.Slides.Add(index, 12)
        title = str(args.get("title", "New Slide"))
        _add_textbox(slide, title, {"name": "Title", "left": 60, "top": 40, "width": 600, "height": 60})
        return {"op": "slides_live_add_slide", "slide_index": index, "slides": int(presentation.Slides.Count)}

    def _slides_live_delete_slide(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        index = int(args["slide_index"])
        _slide(presentation, index).Delete()
        return {"op": "slides_live_delete_slide", "slide_index": index, "slides": int(presentation.Slides.Count)}

    def _slides_live_set_title(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        index = int(args.get("slide_index", 1))
        slide = _slide(presentation, index)
        shape = _shape(slide, {"shape_name": args.get("shape_name", "Title")}, required=False)
        if shape is None:
            shape = _add_textbox(slide, "", {"name": args.get("shape_name", "Title"), "left": 60, "top": 40, "width": 600, "height": 60})
        shape.TextFrame.TextRange.Text = str(args["title"])
        return {"op": "slides_live_set_title", "slide_index": index, "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_add_textbox(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _add_textbox(slide, str(args["text"]), args)
        return {"op": "slides_live_add_textbox", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_format_text(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        font = shape.TextFrame.TextRange.Font
        if args.get("font_name") is not None:
            font.Name = args["font_name"]
        if args.get("font_size") is not None:
            font.Size = float(args["font_size"])
        if args.get("bold") is not None:
            font.Bold = bool(args["bold"])
        if args.get("italic") is not None:
            font.Italic = bool(args["italic"])
        if args.get("color") is not None:
            font.Color.RGB = _rgb(args["color"])
        return {"op": "slides_live_format_text", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_add_shape(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape_type = {"rectangle": 1, "rounded_rectangle": 5, "ellipse": 9}.get(str(args.get("shape_type", "rectangle")), 1)
        shape = slide.Shapes.AddShape(shape_type, float(args.get("left", 60)), float(args.get("top", 140)), float(args.get("width", 160)), float(args.get("height", 80)))
        if args.get("name"):
            shape.Name = args["name"]
        if args.get("fill_color"):
            shape.Fill.ForeColor.RGB = _rgb(args["fill_color"])
        if args.get("line_color"):
            shape.Line.ForeColor.RGB = _rgb(args["line_color"])
        return {"op": "slides_live_add_shape", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_add_image(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        image_path = Path(str(args["image_path"])).resolve()
        if not image_path.is_file():
            raise ValidationError("image_path does not exist", details={"image_path": str(image_path)})
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = slide.Shapes.AddPicture(str(image_path), False, True, float(args.get("left", 60)), float(args.get("top", 240)), float(args.get("width", 80)), float(args.get("height", 80)))
        if args.get("name"):
            shape.Name = args["name"]
        return {"op": "slides_live_add_image", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_set_background(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        index = int(args.get("slide_index", 1))
        slide = _slide(presentation, index)
        slide.FollowMasterBackground = False
        slide.Background.Fill.ForeColor.RGB = _rgb(str(args["color"]))
        return {"op": "slides_live_set_background", "slide_index": index, "color": str(args["color"])}

    def _slides_live_duplicate_slide(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        index = int(args["slide_index"])
        duplicate = _slide(presentation, index).Duplicate()
        return {"op": "slides_live_duplicate_slide", "slide_index": index, "slides": int(presentation.Slides.Count), "duplicate_id": str(_safe_value(duplicate.Item(1), "SlideID", ""))}

    def _slides_live_reorder_slide(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args["slide_index"]))
        slide_id = int(_safe_value(slide, "SlideID", 0))
        to_index = int(args["to_index"])
        slide.MoveTo(to_index)
        return {"op": "slides_live_reorder_slide", "from_index": int(args["slide_index"]), "to_index": to_index, "slide_id": slide_id}

    def _slides_live_set_notes(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        notes = slide.NotesPage
        body = None
        for shape in _iter_collection(notes.Shapes):
            if _shape_text(shape) or int(_safe_value(shape, "Type", 0)) == 14:
                body = shape
                if int(_safe_value(shape, "PlaceholderFormat.Type", 0)) == 2:
                    break
        if body is None:
            body = notes.Shapes.AddTextbox(1, 40, 40, 600, 300)
        body.TextFrame.TextRange.Text = str(args["text"])
        return {"op": "slides_live_set_notes", "slide_index": int(args.get("slide_index", 1)), "text": str(args["text"])}

    def _slides_live_set_slide_size(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        if args.get("width_inches") is not None:
            presentation.PageSetup.SlideWidth = float(args["width_inches"]) * 72
        if args.get("height_inches") is not None:
            presentation.PageSetup.SlideHeight = float(args["height_inches"]) * 72
        return {"op": "slides_live_set_slide_size", "width": float(presentation.PageSetup.SlideWidth), "height": float(presentation.PageSetup.SlideHeight)}

    def _slides_live_save(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        presentation.Save()
        return {"op": "slides_live_save", "saved": bool(_safe_value(presentation, "Saved", True))}

    def _slides_screen_capture(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        output = _absolute_output(str(args["output_path"]))
        try:
            app.Visible = True
        except Exception:
            pass
        window = None
        try:
            window = presentation.Windows(1)
            window.View.GotoSlide(int(args.get("slide_index", 1)))
            window.Activate()
        except Exception:
            window = getattr(app, "ActiveWindow", None)
        if window is None:
            raise ValidationError("PowerPoint presentation has no active window")
        hwnd = _find_powerpoint_window(presentation)
        if not hwnd:
            raise ValidationError("PowerPoint window has no HWND")
        self._screenshotter(hwnd, output)
        return {"op": "slides_screen_capture", "output_path": str(output), "slide_index": int(args.get("slide_index", 1))}

    @staticmethod
    def _get_info(presentation: Any) -> dict[str, Any]:
        page = presentation.PageSetup
        return {
            "name": str(_safe_value(presentation, "Name", "")),
            "path": str(_safe_value(presentation, "FullName", "")),
            "slides": int(_safe_value(presentation.Slides, "Count", 0)),
            "slide_width": float(_safe_value(page, "SlideWidth", 0)),
            "slide_height": float(_safe_value(page, "SlideHeight", 0)),
        }


def _operation_args(operation: dict[str, Any]) -> dict[str, Any]:
    if operation["op"] == "replace_text":
        return {"find_text": operation["find"], "replace_text": operation["replace"], "replace_all": operation.get("occurrence", "all") == "all"}
    return dict(operation.get("args", {}))


def _default_powerpoint_app() -> Any:
    try:
        from win32com.client import Dispatch, GetActiveObject
    except ImportError as exc:
        raise ImportError("pywin32 is required for --backend com") from exc
    try:
        return GetActiveObject("PowerPoint.Application")
    except Exception:
        return Dispatch("PowerPoint.Application")


def _find_presentation(app: Any, target: str) -> Any:
    wanted = Path(target).resolve()
    for presentation in _iter_collection(app.Presentations):
        source = str(_safe_value(presentation, "FullName", ""))
        if target in {source, str(_safe_value(presentation, "Name", ""))} or (source and Path(source).resolve() == wanted):
            return presentation
    raise ValidationError("target PowerPoint presentation is not open", details={"target": target})


def _find_or_open_for_inspect(app: Any, target: Path) -> tuple[Any, bool]:
    try:
        return _find_presentation(app, str(target)), False
    except ValidationError:
        if not target.is_file():
            raise
        try:
            presentation = app.Presentations.Open(
                FileName=str(target),
                ReadOnly=True,
                Untitled=False,
                WithWindow=False,
            )
        except Exception as exc:
            raise ValidationError(
                "PowerPoint could not open the target for inspection",
                details={"target": str(target), "reason": str(exc)},
            ) from exc
        return presentation, True


def _render_slides(presentation: Any, render_dir: Path) -> dict[str, Any]:
    output_dir = render_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    width, height = 1280, 720
    files: list[str] = []
    count = int(_safe_value(presentation.Slides, "Count", 0))
    outputs = [(index, output_dir / f"slide-{index}.png") for index in range(1, count + 1)]
    existing = [str(output) for _, output in outputs if output.exists()]
    if existing:
        raise ValidationError("render output already exists", details={"paths": existing})
    for index, output in outputs:
        try:
            _slide(presentation, index).Export(str(output), "PNG", width, height)
        except Exception as exc:
            raise ValidationError(
                "PowerPoint could not export a slide",
                details={"slide_index": index, "path": str(output), "reason": str(exc)},
            ) from exc
        if not output.is_file():
            raise ValidationError("PowerPoint did not create the rendered slide", details={"slide_index": index, "path": str(output)})
        files.append(str(output))
    return {"directory": str(output_dir), "format": "png", "width": width, "height": height, "files": files}


def _find_powerpoint_window(presentation: Any) -> int:
    try:
        import win32gui
    except ImportError as exc:
        raise BackendUnavailableError("powerpoint-screen-capture") from exc
    wanted = str(_safe_value(presentation, "Name", "")).lower()
    candidates = []

    def collect(hwnd: int, _extra: Any) -> None:
        if win32gui.GetClassName(hwnd) != "PPTFrameClass":
            return
        candidates.append((hwnd, win32gui.GetWindowText(hwnd)))

    win32gui.EnumWindows(collect, None)
    for hwnd, title in candidates:
        if wanted and wanted in title.lower():
            return int(hwnd)
    return int(candidates[0][0]) if candidates else 0


def _slide(presentation: Any, index: int) -> Any:
    count = int(presentation.Slides.Count)
    if index < 1 or index > count:
        raise ValidationError("slide_index is out of range", details={"slide_index": index, "slides": count})
    return presentation.Slides(index)


def _shape(slide: Any, args: dict[str, Any], *, required: bool = True) -> Any | None:
    name = args.get("shape_name")
    index = args.get("shape_index")
    if name:
        for shape in _iter_collection(slide.Shapes):
            if str(_safe_value(shape, "Name", "")) == str(name):
                return shape
    elif index is not None:
        try:
            return slide.Shapes(int(index))
        except Exception:
            pass
    if required:
        raise ValidationError("shape_name or shape_index must identify an existing shape")
    return None


def _add_textbox(slide: Any, text: str, args: dict[str, Any]) -> Any:
    shape = slide.Shapes.AddTextbox(1, float(args.get("left", 60)), float(args.get("top", 40)), float(args.get("width", 600)), float(args.get("height", 60)))
    if args.get("name"):
        shape.Name = args["name"]
    shape.TextFrame.TextRange.Text = text
    return shape


def _slide_text(slide: Any) -> list[dict[str, Any]]:
    values = []
    for index, shape in enumerate(_iter_collection(slide.Shapes), start=1):
        text = _shape_text(shape)
        if text:
            values.append({"shape_index": index, "name": str(_safe_value(shape, "Name", "")), "text": text})
    return values


def _shape_text(shape: Any) -> str:
    try:
        if not bool(shape.HasTextFrame):
            return ""
        return str(shape.TextFrame.TextRange.Text).rstrip("\r\x0b")
    except Exception:
        return ""


def _iter_collection(collection: Any) -> Iterator[Any]:
    if collection is None:
        return iter(())
    try:
        return iter(collection)
    except TypeError:
        count = int(_safe_value(collection, "Count", 0))
        return (collection(index) for index in range(1, count + 1))


def _safe_value(obj: Any, attr: str, default: Any = None) -> Any:
    try:
        value = obj
        for part in attr.split("."):
            value = getattr(value, part)
        if callable(value):
            value = value()
        return value
    except Exception:
        return default


def _absolute_output(value: str) -> Path:
    output = Path(value).resolve()
    if not output.is_absolute():
        raise ValidationError("output_path must be absolute")
    return output


def _rgb(value: str) -> int:
    color = str(value).lstrip("#")
    if len(color) != 6 or any(character not in "0123456789abcdefABCDEF" for character in color):
        raise ValidationError("color must be a six-digit RGB hex value")
    return int(color[0:2], 16) + (int(color[2:4], 16) << 8) + (int(color[4:6], 16) << 16)
