from __future__ import annotations

from pathlib import Path

from whitecollar.adapters.slides import PowerPointComAdapter, _shape


class FakePlaceholder:
    def __init__(self, name: str, placeholder_type: int) -> None:
        self.Name = name
        self.Type = 14
        self.PlaceholderFormat = type("PlaceholderFormat", (), {"Type": placeholder_type})()


class FakePlaceholderSlide:
    def __init__(self) -> None:
        self.Shapes = [
            FakePlaceholder("Title 1", 1),
            FakePlaceholder("Content Placeholder 2", 2),
        ]


class FakeSlide:
    def __init__(self, index: int) -> None:
        self.index = index
        self.exports: list[tuple[str, str, int, int]] = []

    def Export(self, path: str, filter_name: str, width: int, height: int) -> None:
        self.exports.append((path, filter_name, width, height))
        Path(path).write_bytes(b"fake-png")


class FakeSlides:
    def __init__(self, count: int) -> None:
        self._items = [FakeSlide(index) for index in range(1, count + 1)]
        self.Count = count

    def __call__(self, index: int) -> FakeSlide:
        return self._items[index - 1]

    def __iter__(self):
        return iter(self._items)


class FakePresentation:
    Name = "deck.pptx"
    FullName = "C:\\work\\deck.pptx"
    PageSetup = type("PageSetup", (), {"SlideWidth": 960.0, "SlideHeight": 540.0})()

    def __init__(self) -> None:
        self.Slides = FakeSlides(2)
        self.closed = False

    def Close(self) -> None:
        self.closed = True


class FakePresentations:
    def __init__(self, presentation: FakePresentation) -> None:
        self._items = [presentation]

    def __iter__(self):
        return iter(self._items)


class FakePowerPoint:
    def __init__(self, presentation: FakePresentation) -> None:
        self.Presentations = FakePresentations(presentation)


def test_powerpoint_inspect_renders_each_slide(tmp_path):
    target = tmp_path / "deck.pptx"
    target.write_bytes(b"deck")
    presentation = FakePresentation()
    presentation.FullName = str(target)
    adapter = PowerPointComAdapter(app_factory=lambda: FakePowerPoint(presentation))

    value = adapter.inspect(target, render_dir=tmp_path / "renders")

    assert value["backend"] == "powerpoint-com"
    assert value["renders"]["format"] == "png"
    assert value["renders"]["files"] == [
        str((tmp_path / "renders" / "slide-1.png").resolve()),
        str((tmp_path / "renders" / "slide-2.png").resolve()),
    ]
    for slide in presentation.Slides:
        assert slide.exports == [(str((tmp_path / "renders" / f"slide-{slide.index}.png").resolve()), "PNG", 1280, 720)]
    assert not presentation.closed


def test_semantic_title_and_body_select_native_placeholders():
    slide = FakePlaceholderSlide()

    assert _shape(slide, {"shape_name": "Title"}).Name == "Title 1"
    assert _shape(slide, {"shape_name": "Body"}).Name == "Content Placeholder 2"


def test_powerpoint_inspect_does_not_overwrite_render(tmp_path):
    target = tmp_path / "deck.pptx"
    target.write_bytes(b"deck")
    render_dir = tmp_path / "renders"
    render_dir.mkdir()
    (render_dir / "slide-1.png").write_bytes(b"existing")
    presentation = FakePresentation()
    presentation.FullName = str(target)
    adapter = PowerPointComAdapter(app_factory=lambda: FakePowerPoint(presentation))

    try:
        adapter.inspect(target, render_dir=render_dir)
    except Exception as exc:
        assert getattr(exc, "code", None) == "validation_error"
    else:
        raise AssertionError("expected existing render to be rejected")
