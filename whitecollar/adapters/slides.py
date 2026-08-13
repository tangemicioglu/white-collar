from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any, Callable, Iterator

from ..errors import BackendUnavailableError, ValidationError
from ..models import Plan
from ..office_capture import capture_window
from ..slides_ops import SLIDES_COM_MUTATING_OPERATIONS, SLIDES_COM_OPERATIONS

SLIDES_COM_SELF_WRITING_OPERATIONS = {"slides_live_save_template", "slides_live_export_pdf"}


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
        if any(operation["op"] == "slides_live_create_presentation" for operation in plan.operations):
            return self._slides_live_create_presentation(plan, dry_run=dry_run)
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
            args = _operation_args(operation)
            if name in SLIDES_COM_SELF_WRITING_OPERATIONS and plan.write.path:
                args.setdefault("output_path", plan.write.path)
            if name not in SLIDES_COM_OPERATIONS and name not in {"replace_text"}:
                raise ValidationError(f"unsupported PowerPoint COM operation: {name}")
            if dry_run and name in SLIDES_COM_MUTATING_OPERATIONS:
                operations.append({"op": name, "dry_run": True, "args": args})
                continue
            dispatch_name = "slides_live_replace_text" if name == "replace_text" else name
            method = getattr(self, f"_{dispatch_name}", None)
            if method is None:
                raise ValidationError(f"PowerPoint operation is registered but not implemented: {name}")
            operations.append(method(app, presentation, args))
        if not dry_run and plan.write.mode != "none" and not any(
            operation["op"] in SLIDES_COM_SELF_WRITING_OPERATIONS for operation in plan.operations
        ):
            self._commit_write(app, presentation, plan)
        return {"backend": "powerpoint-com", "written": not dry_run, "operations": operations}

    def _slides_live_create_presentation(self, plan: Plan, *, dry_run: bool) -> dict[str, Any]:
        if len(plan.operations) != 1 or plan.operations[0]["op"] != "slides_live_create_presentation":
            raise ValidationError("slides_live_create_presentation must be a standalone plan")
        target = Path(plan.target.path)
        if target.suffix.lower() != ".pptx":
            raise ValidationError("PowerPoint creation output must use the .pptx extension", details={"target": str(target)})
        if target.exists():
            raise ValidationError("creation output already exists", details={"target": str(target)})
        if dry_run:
            return {
                "backend": "powerpoint-com",
                "written": False,
                "operations": [{"op": "slides_live_create_presentation", "dry_run": True, "args": {}}],
            }
        app = self._get_app()
        presentation = None
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            presentation = app.Presentations.Add(True)
            presentation.Slides.Add(1, 12)  # ppLayoutBlank: a usable blank first slide.
            presentation.SaveAs(FileName=str(target), FileFormat=24)  # ppSaveAsOpenXMLPresentation.
            if not target.is_file():
                raise ValidationError("PowerPoint did not create the requested presentation", details={"target": str(target)})
            operation = {
                "op": "slides_live_create_presentation",
                "created": True,
                "path": str(target),
                "slides": int(_safe_value(presentation.Slides, "Count", 0)),
            }
            return {"backend": "powerpoint-com", "written": True, "operations": [operation]}
        except ValidationError:
            raise
        except Exception as exc:
            raise ValidationError(
                "PowerPoint could not create the requested presentation",
                details={"target": str(target), "reason": str(exc)},
            ) from exc
        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception:
                    pass

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

    def _slides_live_get_masters(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        masters = []
        try:
            designs = presentation.Designs
        except Exception:
            designs = []
        for index, design in enumerate(_iter_collection(designs), start=1):
            master = _safe_value(design, "SlideMaster")
            masters.append({"index": index, "name": str(_safe_value(design, "Name", "")), "master_name": str(_safe_value(master, "Name", ""))})
        if not masters:
            try:
                master = presentation.SlideMaster
            except Exception:
                master = None
            if master is None and int(_safe_value(presentation.Slides, "Count", 0)):
                master = _safe_value(presentation.Slides(1), "Master")
            if master is not None:
                masters.append({"index": 1, "name": "", "master_name": str(_safe_value(master, "Name", ""))})
        return {"op": "slides_live_get_masters", "masters": masters, "count": len(masters)}

    def _slides_live_get_layouts(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        master = _presentation_master(presentation, args)
        layouts = []
        for index, layout in enumerate(_iter_collection(master.CustomLayouts), start=1):
            layouts.append({"index": index, "name": str(_safe_value(layout, "Name", "")), "type": _safe_value(layout, "Type")})
        return {"op": "slides_live_get_layouts", "layouts": layouts, "count": len(layouts)}

    def _slides_live_get_placeholders(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        if args.get("slide_index") is not None:
            shapes = _slide(presentation, int(args["slide_index"])).Shapes
        else:
            shapes = _presentation_master(presentation, args).Shapes
        placeholders = []
        for index, shape in enumerate(_iter_collection(shapes), start=1):
            if int(_safe_value(shape, "Type", 0)) != 14:
                continue
            placeholders.append({
                "index": index,
                "name": str(_safe_value(shape, "Name", "")),
                "type": _safe_value(shape, "PlaceholderFormat.Type"),
                "text": _shape_text(shape),
            })
        return {"op": "slides_live_get_placeholders", "placeholders": placeholders, "count": len(placeholders)}

    def _slides_live_get_notes(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        indices = [int(args["slide_index"])] if args.get("slide_index") is not None else list(range(1, int(presentation.Slides.Count) + 1))
        notes = []
        for index in indices:
            slide = _slide(presentation, index)
            text = []
            for shape in _iter_collection(slide.NotesPage.Shapes):
                value = _shape_text(shape)
                if value and int(_safe_value(shape, "PlaceholderFormat.Type", 0)) in {2, 3}:
                    text.append(value)
            notes.append({"slide_index": index, "text": "\n".join(text)})
        return {"op": "slides_live_get_notes", "notes": notes, "count": len(notes)}

    def _slides_live_get_sections(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        try:
            sections = presentation.SectionProperties
        except Exception as exc:
            raise ValidationError("PowerPoint presentation has no section properties") from exc
        values = []
        for index in range(1, int(_safe_value(sections, "Count", 0)) + 1):
            values.append({
                "index": index,
                "name": str(sections.Name(index)),
                "first_slide": int(sections.FirstSlide(index)),
                "slides": int(sections.SlidesCount(index)),
            })
        return {"op": "slides_live_get_sections", "sections": values, "count": len(values)}

    def _slides_live_get_media(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        values = []
        for slide_index, slide in enumerate(_iter_collection(presentation.Slides), start=1):
            for shape_index, shape in enumerate(_iter_collection(slide.Shapes), start=1):
                shape_type = int(_safe_value(shape, "Type", 0))
                media_type = _safe_value(shape, "MediaType")
                if media_type is None and shape_type != 16:
                    continue
                values.append({
                    "slide_index": slide_index,
                    "shape_index": shape_index,
                    "name": str(_safe_value(shape, "Name", "")),
                    "shape_type": shape_type,
                    "media_type": media_type,
                    "path": str(_safe_value(shape, "LinkFormat.SourceFullName", "") or ""),
                })
        return {"op": "slides_live_get_media", "media": values, "count": len(values)}

    def _slides_live_apply_template(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        source = Path(str(args["source_path"])).resolve()
        if not source.is_file():
            raise ValidationError("source_path does not exist", details={"source_path": str(source)})
        presentation.ApplyTemplate(str(source))
        return {"op": "slides_live_apply_template", "source_path": str(source), "applied": True}

    def _slides_live_save_template(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        output = _absolute_output(str(args.get("output_path", "")))
        if output.suffix.lower() != ".potx":
            raise ValidationError("template output must use the .potx extension")
        if output.exists():
            raise ValidationError("template output already exists", details={"path": str(output)})
        output.parent.mkdir(parents=True, exist_ok=True)
        saved_copy = getattr(presentation, "SaveCopyAs", None)
        try:
            if callable(saved_copy):
                saved_copy(FileName=str(output), FileFormat=26)  # ppSaveAsOpenXMLTemplate
            else:
                raise RuntimeError("SaveCopyAs is unavailable")
        except Exception:
            # Some Office builds expose SaveCopyAs but reject a format change.
            # SaveAs is reliable, but changes the open presentation's identity;
            # restore the original open file before returning to the caller.
            source = Path(str(_safe_value(presentation, "FullName", "")))
            presentation.SaveAs(FileName=str(output), FileFormat=26)  # ppSaveAsOpenXMLTemplate
            if source.resolve() != output.resolve():
                presentation.Close()
                app.Presentations.Open(FileName=str(source), ReadOnly=False, Untitled=False, WithWindow=True)
        if not output.is_file():
            raise ValidationError("PowerPoint did not create the requested template", details={"path": str(output)})
        return {"op": "slides_live_save_template", "path": str(output), "bytes": output.stat().st_size}

    def _slides_live_set_layout(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        layout = _find_layout(_presentation_master(presentation, args), args["layout"])
        slide.CustomLayout = layout
        return {"op": "slides_live_set_layout", "slide_index": int(args.get("slide_index", 1)), "layout": str(_safe_value(layout, "Name", args["layout"]))}

    def _slides_live_group(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        group = _shape_range(slide, args).Group()
        return {"op": "slides_live_group", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(group, "Name", ""))}

    def _slides_live_ungroup(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        shape.Ungroup()
        return {"op": "slides_live_ungroup", "slide_index": int(args.get("slide_index", 1)), "ungrouped": True}

    def _slides_live_align(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        command = _enum(args["alignment"], {"left": 0, "center": 1, "right": 2, "top": 3, "middle": 4, "bottom": 5}, "alignment")
        _shape_range(slide, args).Align(command, bool(args.get("relative_to_slide", True)))
        return {"op": "slides_live_align", "alignment": args["alignment"]}

    def _slides_live_distribute(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        direction = _enum(args["direction"], {"horizontal": 0, "vertical": 1}, "direction")
        _shape_range(slide, args).Distribute(direction, bool(args.get("relative_to_slide", True)))
        return {"op": "slides_live_distribute", "direction": args["direction"]}

    def _slides_live_z_order(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        command = _enum(args["command"], {"bring_to_front": 0, "send_to_back": 1, "bring_forward": 2, "send_backward": 3}, "command")
        shape = _shape(slide, args)
        shape.ZOrder(command)
        return {"op": "slides_live_z_order", "command": args["command"], "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_crop_image(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        try:
            picture = shape.PictureFormat
        except Exception:
            raise ValidationError("shape does not expose PictureFormat")
        for name in ("CropLeft", "CropTop", "CropRight", "CropBottom"):
            key = name[4:].lower()
            if key in args:
                setattr(picture, name, float(args[key]))
        return {"op": "slides_live_crop_image", "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_rotate_shape(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        shape.Rotation = float(args["degrees"])
        return {"op": "slides_live_rotate_shape", "degrees": float(args["degrees"]), "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_add_section(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        sections = presentation.SectionProperties
        slide_index = int(args.get("slide_index", 1))
        sections.AddBeforeSlide(slide_index, str(args["name"]))
        return {"op": "slides_live_add_section", "name": str(args["name"]), "count": int(sections.Count)}

    def _slides_live_delete_section(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        sections = presentation.SectionProperties
        index = int(args.get("section_index", 1))
        if index < 1 or index > int(sections.Count):
            raise ValidationError("section_index is out of range")
        sections.Delete(index, False)
        return {"op": "slides_live_delete_section", "section_index": index, "count": int(sections.Count)}

    def _slides_live_set_slide_visibility(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        index = int(args.get("slide_index", 1))
        slide = _slide(presentation, index)
        slide.SlideShowTransition.Hidden = not bool(args["visible"])
        return {"op": "slides_live_set_slide_visibility", "slide_index": index, "visible": bool(args["visible"])}

    def _slides_live_set_slide_numbers(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        visible = bool(args["visible"])
        changed = 0
        errors = []
        if args.get("slide_index") is not None:
            slides = [_slide(presentation, int(args["slide_index"]))]
        else:
            slides = list(_iter_collection(presentation.Slides))
        for slide in slides:
            try:
                slide.HeadersFooters.SlideNumber.Visible = visible
                changed += 1
            except Exception as exc:
                errors.append(str(exc))
        try:
            presentation.SlideMaster.HeadersFooters.SlideNumber.Visible = visible
        except Exception as exc:
            errors.append(str(exc))
        if changed == 0:
            raise ValidationError(
                "PowerPoint did not expose a writable slide-number footer",
                details={"visible": visible, "reason": errors[0] if errors else "unknown error"},
            )
        return {"op": "slides_live_set_slide_numbers", "visible": visible, "slides": changed}

    def _slides_live_add_table(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        rows, columns = int(args["rows"]), int(args["columns"])
        shape = slide.Shapes.AddTable(rows, columns, float(args.get("left", 60)), float(args.get("top", 140)), float(args.get("width", 600)), float(args.get("height", 220)))
        if args.get("name"):
            shape.Name = str(args["name"])
        data = args.get("data", [])
        for row_index, row in enumerate(data, start=1):
            if row_index > rows or not isinstance(row, list):
                break
            for column_index, value in enumerate(row, start=1):
                if column_index <= columns:
                    shape.Table.Cell(row_index, column_index).Shape.TextFrame.TextRange.Text = str(value)
        return {"op": "slides_live_add_table", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", "")), "rows": rows, "columns": columns}

    def _slides_live_set_table_cell(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        try:
            table = shape.Table
        except Exception:
            raise ValidationError("shape is not a table")
        table.Cell(int(args["row"]), int(args["column"])).Shape.TextFrame.TextRange.Text = str(args["text"])
        return {"op": "slides_live_set_table_cell", "row": int(args["row"]), "column": int(args["column"]), "text": str(args["text"])}

    def _slides_live_add_chart(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        chart_type = _chart_type(args["chart_type"])
        shape = slide.Shapes.AddChart2(201, chart_type, float(args.get("left", 60)), float(args.get("top", 140)), float(args.get("width", 600)), float(args.get("height", 320)))
        if args.get("name"):
            shape.Name = str(args["name"])
        chart = shape.Chart
        if args.get("title") is not None:
            chart.HasTitle = True
            chart.ChartTitle.Text = str(args["title"])
        if args.get("data") is not None:
            _write_chart_data(chart, args["data"])
        return {"op": "slides_live_add_chart", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", "")), "chart_type": chart_type}

    def _slides_live_add_smartart(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        try:
            layouts = app.SmartArtLayouts
        except Exception as exc:
            raise ValidationError("PowerPoint does not expose SmartArt layouts") from exc
        layout = _find_collection_item(layouts, args.get("layout", 1))
        shape = slide.Shapes.AddSmartArt(layout, float(args.get("left", 60)), float(args.get("top", 140)), float(args.get("width", 600)), float(args.get("height", 260)))
        if args.get("name"):
            shape.Name = str(args["name"])
        if args.get("nodes") is not None:
            try:
                nodes = shape.SmartArt.AllNodes
                count = int(nodes.Count)
            except Exception as exc:
                raise ValidationError("PowerPoint did not expose SmartArt nodes") from exc
            values = list(args["nodes"])
            if len(values) > count:
                raise ValidationError(
                    "nodes contains more entries than the selected SmartArt layout",
                    details={"requested": len(values), "available": count},
                )
            for index, value in enumerate(values, start=1):
                nodes(index).TextFrame2.TextRange.Text = str(value)
        return {"op": "slides_live_add_smartart", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", ""))}

    def _slides_live_add_media(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        media = Path(str(args["media_path"])).resolve()
        if not media.is_file():
            raise ValidationError("media_path does not exist", details={"media_path": str(media)})
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = slide.Shapes.AddMediaObject2(
            FileName=str(media), LinkToFile=False, SaveWithDocument=True,
            Left=float(args.get("left", 60)), Top=float(args.get("top", 140)),
            Width=float(args.get("width", -1)), Height=float(args.get("height", -1)),
        )
        if args.get("name"):
            shape.Name = str(args["name"])
        return {"op": "slides_live_add_media", "slide_index": int(args.get("slide_index", 1)), "shape": str(_safe_value(shape, "Name", "")), "path": str(media)}

    def _slides_live_set_hyperlink(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        hyperlink = shape.ActionSettings(1).Hyperlink
        hyperlink.Address = str(args["url"])
        if args.get("sub_address") is not None:
            hyperlink.SubAddress = str(args["sub_address"])
        return {"op": "slides_live_set_hyperlink", "shape": str(_safe_value(shape, "Name", "")), "url": str(args["url"])}

    def _slides_live_set_alt_text(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        shape.AlternativeText = str(args["text"])
        return {"op": "slides_live_set_alt_text", "shape": str(_safe_value(shape, "Name", "")), "text": str(args["text"])}

    def _slides_live_set_transition(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        transition = slide.SlideShowTransition
        transition.EntryEffect = _transition_effect(args["effect"])
        if args.get("advance_on_click") is not None:
            transition.AdvanceOnClick = bool(args["advance_on_click"])
        if args.get("advance_seconds") is not None:
            transition.AdvanceOnTime = True
            transition.AdvanceTime = float(args["advance_seconds"])
        return {"op": "slides_live_set_transition", "slide_index": int(args.get("slide_index", 1)), "effect": args["effect"]}

    def _slides_live_add_animation(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        slide = _slide(presentation, int(args.get("slide_index", 1)))
        shape = _shape(slide, args)
        effect = _animation_effect(args["effect"])
        trigger = _enum(args.get("trigger", "on_click"), {"on_click": 1, "with_previous": 2, "after_previous": 3}, "trigger")
        sequence = slide.TimeLine.MainSequence
        sequence.AddEffect(shape, effect, 0, trigger)
        return {"op": "slides_live_add_animation", "slide_index": int(args.get("slide_index", 1)), "effect": args["effect"], "trigger": args.get("trigger", "on_click")}

    def _slides_live_export_pdf(self, app: Any, presentation: Any, args: dict[str, Any]) -> dict[str, Any]:
        output = _absolute_output(str(args.get("output_path", "")))
        if output.exists():
            raise ValidationError("PDF output already exists", details={"path": str(output)})
        output.parent.mkdir(parents=True, exist_ok=True)
        saved_copy = getattr(presentation, "SaveCopyAs", None)
        if not callable(saved_copy):
            raise ValidationError("PowerPoint does not expose SaveCopyAs for PDF export")
        saved_copy(FileName=str(output), FileFormat=32)  # ppSaveAsPDF
        if not output.is_file():
            raise ValidationError("PowerPoint did not create the requested PDF", details={"path": str(output)})
        return {"op": "slides_live_export_pdf", "path": str(output), "bytes": output.stat().st_size}

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


def _shape_range(slide: Any, args: dict[str, Any]) -> Any:
    names = args.get("shape_names")
    indices = args.get("shape_indices")
    if names is None and indices is None:
        shape = _shape(slide, args)
        item = str(_safe_value(shape, "Name", "")) if args.get("shape_name") else int(args.get("shape_index", 1))
        return slide.Shapes.Range((item,))
    items = tuple(str(value) for value in names) if names is not None else tuple(int(value) for value in indices)
    if len(items) < 1:
        raise ValidationError("shape_names or shape_indices must contain at least one item")
    try:
        return slide.Shapes.Range(items)
    except Exception:
        try:
            return slide.Shapes.Range(list(items))
        except Exception as exc:
            raise ValidationError("PowerPoint could not resolve the requested shape range", details={"items": list(items)}) from exc


def _presentation_master(presentation: Any, args: dict[str, Any]) -> Any:
    try:
        designs = presentation.Designs
    except Exception:
        designs = []
    selector = args.get("master")
    if selector is not None:
        return _find_collection_item(designs, selector).SlideMaster
    try:
        master = presentation.SlideMaster
    except Exception:
        master = None
    if master is None and int(_safe_value(presentation.Slides, "Count", 0)):
        master = _safe_value(presentation.Slides(1), "Master")
    if master is None:
        raise ValidationError("PowerPoint presentation has no slide master")
    return master


def _find_layout(master: Any, selector: Any) -> Any:
    return _find_collection_item(master.CustomLayouts, selector)


def _find_collection_item(collection: Any, selector: Any) -> Any:
    if isinstance(selector, bool):
        raise ValidationError("collection selector must be a name or positive integer")
    if isinstance(selector, int) or (isinstance(selector, str) and selector.isdigit()):
        index = int(selector)
        if index < 1 or index > int(_safe_value(collection, "Count", 0)):
            raise ValidationError("collection index is out of range", details={"index": index})
        return collection(index)
    wanted = str(selector).casefold()
    for item in _iter_collection(collection):
        if str(_safe_value(item, "Name", "")).casefold() == wanted:
            return item
    raise ValidationError("named PowerPoint collection item was not found", details={"name": str(selector)})


def _chart_type(value: Any) -> int:
    if isinstance(value, bool):
        raise ValidationError("chart_type must be a name or integer")
    if isinstance(value, int) or (isinstance(value, str) and value.isdigit()):
        return int(value)
    values = {"column": 51, "bar": 57, "line": 4, "pie": 5, "area": 1, "scatter": -4169}
    key = str(value).casefold()
    if key not in values:
        raise ValidationError("unknown chart_type", details={"allowed": sorted(values)})
    return values[key]


def _write_chart_data(chart: Any, data: Any) -> None:
    """Write a bounded rectangular matrix into an embedded PowerPoint chart."""
    if not isinstance(data, list) or not data or not all(isinstance(row, list) and row for row in data):
        raise ValidationError("chart data must be a non-empty array of non-empty rows")
    width = len(data[0])
    if any(len(row) != width for row in data):
        raise ValidationError("chart data must be rectangular")
    workbook = None
    try:
        chart.ChartData.Activate()
        workbook = chart.ChartData.Workbook
        sheet = workbook.Worksheets(1)
        for row_index, row in enumerate(data, start=1):
            for column_index, value in enumerate(row, start=1):
                sheet.Cells(row_index, column_index).Value = value
        address = sheet.Range(
            sheet.Cells(1, 1),
            sheet.Cells(len(data), width),
        ).Address
        # PowerPoint requires the embedded worksheet name in this address;
        # the bare Excel range returned by Range.Address is rejected by some
        # Office builds even though the cells themselves were written.
        source = f"'{sheet.Name}'!{address}"
        chart.SetSourceData(Source=source, PlotBy=2)
    except Exception as exc:
        raise ValidationError("PowerPoint could not write the chart data", details={"reason": str(exc)}) from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close()
            except Exception:
                pass


def _transition_effect(value: Any) -> int:
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    values = {"none": 257, "cut": 257, "fade": 3845, "appear": 3844, "wipe_right": 2817, "wipe_left": 2819}
    key = str(value).casefold()
    if key not in values:
        raise ValidationError("unknown transition effect", details={"allowed": sorted(values)})
    return values[key]


def _animation_effect(value: Any) -> int:
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    values = {"appear": 1, "fly": 2, "fade": 10, "zoom": 23}
    key = str(value).casefold()
    if key not in values:
        raise ValidationError("unknown animation effect", details={"allowed": sorted(values)})
    return values[key]


def _enum(value: Any, values: dict[str, int], name: str) -> int:
    key = str(value).casefold()
    if key not in values:
        raise ValidationError(f"invalid {name}: {value}", details={"allowed": sorted(values)})
    return values[key]


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
