from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from whitecollar.adapters import word_com
from whitecollar.adapters.word_com import Win32WordComAdapter
from whitecollar.errors import ValidationError


class Collection:
    def __init__(self, *items):
        self.items = list(items)
        self.Count = len(self.items)

    def __iter__(self):
        return iter(self.items)

    def __call__(self, index):
        return self.items[index - 1]


class FakeDocument:
    Name = "brief.docx"
    TrackRevisions = False
    Saved = True

    def __init__(self, path: Path) -> None:
        self.FullName = str(path)
        paragraph = SimpleNamespace(Range=SimpleNamespace(Text="Rendered text\r"), Style=SimpleNamespace(NameLocal="Normal"))
        self.Paragraphs = Collection(paragraph)
        self.Sections = Collection(SimpleNamespace())
        self.export_args = None
        self.closed = False

    def ComputeStatistics(self, _kind):
        return 2

    def ExportAsFixedFormat(self, **kwargs):
        self.export_args = kwargs
        Path(kwargs["OutputFileName"]).write_bytes(b"%PDF-fake")

    def Close(self, **_kwargs):
        self.closed = True


class FakeApp:
    def __init__(self, document: FakeDocument) -> None:
        self.Documents = Collection(document)


def test_word_inspect_renders_each_page(tmp_path, monkeypatch):
    target = tmp_path / "brief.docx"
    target.write_bytes(b"docx")
    document = FakeDocument(target)

    def fake_run(command, *, check, capture_output, text):
        prefix = Path(command[-1])
        for index in (1, 2):
            Path(f"{prefix}-{index}.png").write_bytes(b"fake-png")
        return SimpleNamespace(returncode=0, stderr="")

    monkeypatch.setattr(word_com.shutil, "which", lambda name: "pdftoppm.exe")
    monkeypatch.setattr(word_com.subprocess, "run", fake_run)
    adapter = Win32WordComAdapter(app_factory=lambda: FakeApp(document))

    value = adapter.inspect(target, render_dir=tmp_path / "renders")

    assert value["backend"] == "word-com"
    assert value["text"] == [{"index": 1, "text": "Rendered text", "style": "Normal"}]
    assert value["renders"]["pages"] == 2
    assert value["renders"]["files"] == [
        str((tmp_path / "renders" / "page-1.png").resolve()),
        str((tmp_path / "renders" / "page-2.png").resolve()),
    ]
    assert document.export_args["ExportFormat"] == 17
    assert not document.closed


def test_word_inspect_does_not_overwrite_render(tmp_path, monkeypatch):
    target = tmp_path / "brief.docx"
    target.write_bytes(b"docx")
    render_dir = tmp_path / "renders"
    render_dir.mkdir()
    (render_dir / "page-1.png").write_bytes(b"existing")
    document = FakeDocument(target)
    monkeypatch.setattr(word_com.shutil, "which", lambda name: "pdftoppm.exe")
    adapter = Win32WordComAdapter(app_factory=lambda: FakeApp(document))

    with pytest.raises(ValidationError, match="already exists"):
        adapter.inspect(target, render_dir=render_dir)
    assert document.export_args is None
