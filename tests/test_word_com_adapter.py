from __future__ import annotations

from types import SimpleNamespace

import pytest

from whitecollar.adapters.word_com import Win32WordComAdapter
from whitecollar.errors import ValidationError
from whitecollar.models import Plan
from whitecollar.word_ops import WORD_COM_OPERATIONS


class Collection:
    def __init__(self, *items):
        self.items = list(items)
        self.Count = len(self.items)

    def __iter__(self):
        return iter(self.items)

    def __call__(self, index):
        return self.items[index - 1]


class FakeFind:
    def __init__(self):
        self.Replacement = SimpleNamespace(ClearFormatting=lambda: None, Text="")
        self.calls = []

    def ClearFormatting(self):
        pass

    def Execute(self, *args, **kwargs):
        self.calls.append(kwargs if kwargs else args)
        return 2 if kwargs else len(self.calls) == 1


class FakeRange:
    def __init__(self, text="Draft"):
        self.Text = text
        self.Start = 0
        self.End = len(text)
        self.Find = FakeFind()
        self.Replacement = SimpleNamespace(ClearFormatting=lambda: None, Text="")
        self.Font = SimpleNamespace()
        self.Paragraphs = Collection()

    @property
    def Duplicate(self):
        return self

    def Collapse(self, direction):
        pass

    def InsertAfter(self, text):
        self.Text += text

    def InsertBefore(self, text):
        self.Text = text + self.Text

    def Delete(self):
        self.Text = ""


class FakeParagraph:
    def __init__(self, text="Draft"):
        self.Range = FakeRange(text)
        self.Format = SimpleNamespace()
        self.Style = SimpleNamespace(NameLocal="Normal")


class FakeDoc:
    def __init__(self):
        self.Name = "brief.docx"
        self.FullName = r"C:\work\brief.docx"
        self.TrackRevisions = False
        self.Saved = True
        self.Content = FakeRange("Draft")
        self._find = self.Content.Find
        self.Paragraphs = Collection(FakeParagraph("Draft"))
        self.Sections = Collection()
        self.Tables = Collection()
        self.Bookmarks = SimpleNamespace(Exists=lambda name: False)
        self.Application = SimpleNamespace(Selection=SimpleNamespace(Range=FakeRange()))

    def Range(self, start, end):
        result = FakeRange("Draft" if int(start) == 0 else "")
        result.Start = int(start)
        result.End = int(end)
        result.Find = self._find
        return result

    def SaveCopyAs(self, path):
        self.saved_copy = path

    def Save(self):
        self.saved = True


class FakeUndoRecord:
    def __init__(self):
        self.events = []

    def StartCustomRecord(self, name):
        self.events.append(("start", name))

    def EndCustomRecord(self):
        self.events.append(("end", None))


class FakeApp:
    def __init__(self):
        self.document = FakeDoc()
        self.Documents = Collection(self.document)
        self.UndoRecord = FakeUndoRecord()


class FakeShapes:
    def __init__(self, texts):
        self.items = []
        for text in texts:
            shape = SimpleNamespace(TextEffect=SimpleNamespace(Text=text))
            shape.Delete = lambda shape=shape: self.items.remove(shape)
            self.items.append(shape)

    @property
    def Count(self):
        return len(self.items)

    def __iter__(self):
        return iter(list(self.items))


class FakeHeaderFooter:
    def __init__(self, texts):
        self.Shapes = FakeShapes(texts)


class FakeWatermarkSection:
    def __init__(self, headers, footers):
        self._headers = {1: FakeHeaderFooter(headers), 2: FakeHeaderFooter(()), 3: FakeHeaderFooter(())}
        self._footers = {1: FakeHeaderFooter(footers), 2: FakeHeaderFooter(()), 3: FakeHeaderFooter(())}

    def Headers(self, kind):
        return self._headers[kind]

    def Footers(self, kind):
        return self._footers[kind]


class FakeWatermarkDoc(FakeDoc):
    def __init__(self):
        super().__init__()
        self.TrackRevisions = True
        self.Sections = Collection(FakeWatermarkSection(("DRAFT", "APPROVED"), ("DRAFT",)))
        self.ActiveWindow = SimpleNamespace(View=SimpleNamespace(SeekView=0))

    def Activate(self):
        self.activated = True


def plan_for(operation, args, *, policy="edit", write=None):
    return Plan.from_dict(
        {
            "schema": "white-collar.plan/v1",
            "app": "word",
            "target": {"path": r"C:\work\brief.docx"},
            "policy": policy,
            "operations": [{"op": operation, "args": args}],
            "write": write or {"mode": "in-place", "snapshot": r"C:\work\before.docx"},
        }
    )


def test_reference_live_vocabulary_has_a_constrained_handler():
    adapter = Win32WordComAdapter(app_factory=lambda: FakeApp())
    unhandled = [name for name in WORD_COM_OPERATIONS if name != "word_live_list_open" and not hasattr(adapter, f"_{name}")]
    assert unhandled == []


def test_replace_routes_through_fake_word_and_groups_undo():
    app = FakeApp()
    adapter = Win32WordComAdapter(app_factory=lambda: app)
    plan = plan_for("word_live_replace_text", {"find_text": "Draft", "replace_text": "Final"})
    value = adapter.apply(plan, dry_run=False)
    assert value["backend"] == "word-com"
    assert value["operations"][0]["op"] == "word_live_replace_text"
    assert app.document.Content.Find.calls == [()]
    assert app.UndoRecord.events[0][0] == "start"
    assert app.UndoRecord.events[-1][0] == "end"


def test_legacy_replace_text_plan_is_supported_by_com_backend():
    app = FakeApp()
    adapter = Win32WordComAdapter(app_factory=lambda: app)
    plan = Plan.from_dict(
        {
            "schema": "white-collar.plan/v1",
            "app": "word",
            "target": {"path": r"C:\work\brief.docx"},
            "policy": "edit",
            "operations": [{"op": "replace_text", "find": "Draft", "replace": "Final"}],
            "write": {"mode": "in-place", "snapshot": r"C:\work\before.docx"},
        }
    )
    value = adapter.apply(plan, dry_run=False)
    assert value["operations"][0]["op"] == "word_live_replace_text"


def test_read_operation_works_with_none_write_and_read_only_policy():
    app = FakeApp()
    adapter = Win32WordComAdapter(app_factory=lambda: app)
    plan = plan_for("word_live_get_text", {}, policy="read-only", write={"mode": "none"})
    value = adapter.apply(plan, dry_run=False)
    assert value["operations"][0]["total_paragraphs"] == 1


def test_mutating_com_dry_run_never_dispatches_to_document():
    app = FakeApp()
    calls = []
    adapter = Win32WordComAdapter(app_factory=lambda: calls.append(app) or app)
    plan = plan_for("word_live_insert_text", {"text": "hello"})
    value = adapter.apply(plan, dry_run=True)
    assert value["written"] is False
    assert value["operations"][0]["dry_run"] is True
    assert calls == []
    assert app.UndoRecord.events == []


def test_remove_watermark_deletes_exact_wordart_matches_and_restores_tracking():
    app = SimpleNamespace(
        document=FakeWatermarkDoc(),
        UndoRecord=FakeUndoRecord(),
    )
    app.Documents = Collection(app.document)
    adapter = Win32WordComAdapter(app_factory=lambda: app)
    plan = plan_for("word_live_remove_watermark", {})

    value = adapter.apply(plan, dry_run=False)

    operation = value["operations"][0]
    assert operation["removed"] == 2
    assert operation["text"] == "DRAFT"
    assert app.document.TrackRevisions is True
    assert app.document.Sections(1).Headers(1).Shapes.Count == 1
    assert app.document.Sections(1).Headers(1).Shapes.items[0].TextEffect.Text == "APPROVED"
    assert app.document.Sections(1).Footers(1).Shapes.Count == 0
    assert app.UndoRecord.events[0][0] == "start"
    assert app.UndoRecord.events[-1][0] == "end"


def test_com_preflight_rejects_existing_save_as_before_dispatching_operation(tmp_path):
    app = FakeApp()
    output = tmp_path / "existing.docx"
    output.write_text("do not overwrite", encoding="utf-8")
    plan = plan_for(
        "word_live_insert_text",
        {"text": "hello"},
        write={"mode": "save-as", "path": str(output)},
    )
    adapter = Win32WordComAdapter(app_factory=lambda: app)
    with pytest.raises(ValidationError, match="already exists"):
        adapter.apply(plan, dry_run=False)
    assert not hasattr(app.document, "saved_copy")


@pytest.mark.parametrize("operation,args", [("word_live_insert_text", {}), ("word_live_find_text", {})])
def test_operation_specific_arguments_are_validated(operation, args):
    with pytest.raises(ValidationError, match="missing required"):
        plan_for(operation, args)
