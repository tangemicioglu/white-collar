from __future__ import annotations

import json

from conftest import write_plan
from whitecollar.adapters.word import OoxmlWordAdapter
from whitecollar.cli import main
from whitecollar.engine import RuntimeAdapters


class FakeSlides:
    def inspect(self, target):
        return {"slides": 7, "title": "Review"}

    def apply(self, plan, *, dry_run):
        return {"changes": [{"operation": 0, "matches": 3}], "written": not dry_run}


class FakeMail:
    def search(self, query, *, limit):
        return [{"id": "m-1", "subject": "Roadmap", "query": query}][:limit]

    def read(self, message_id):
        return {"id": message_id, "subject": "Roadmap", "body": "Ship it"}


def adapters() -> RuntimeAdapters:
    return RuntimeAdapters(OoxmlWordAdapter(), FakeSlides(), FakeMail())


def output(capsys):
    return json.loads(capsys.readouterr().out)


def test_word_inspect_smoke(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    assert main(["word", "inspect", str(target)], adapters=adapters()) == 0
    value = output(capsys)
    assert value["schema"] == "white-collar.result/v1"
    assert value["command"] == "word.inspect"
    assert value["policy"] == "read-only"


def test_word_apply_dry_run_smoke(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    plan = write_plan(tmp_path / "plan.json", target, tmp_path / "final.docx")
    assert main(["word", "apply", "--plan", str(plan), "--dry-run"], adapters=adapters()) == 0
    value = output(capsys)
    assert value["dry_run"] is True
    assert value["changes"][0]["matches"] == 1


def test_slides_commands_use_only_slides_adapter(tmp_path, capsys):
    deck = tmp_path / "deck.pptx"
    assert main(["slides", "inspect", str(deck)], adapters=adapters()) == 0
    assert output(capsys)["data"]["slides"] == 7
    plan = write_plan(tmp_path / "plan.json", deck, tmp_path / "new.pptx", app="slides")
    assert main(["slides", "apply", "--plan", str(plan), "--dry-run"], adapters=adapters()) == 0
    assert output(capsys)["changes"][0]["matches"] == 3


def test_mail_search_and_read_default_to_read_only(capsys):
    assert main(["mail", "search", "--query", "roadmap"], adapters=adapters()) == 0
    searched = output(capsys)
    assert searched["policy"] == "read-only"
    assert searched["data"][0]["id"] == "m-1"
    assert main(["mail", "read", "--id", "m-1"], adapters=adapters()) == 0
    assert output(capsys)["data"]["body"] == "Ship it"


def test_cli_validation_errors_are_json(capsys):
    assert main(["mail", "search", "--query", "x", "--limit", "101"], adapters=adapters()) == 2
    value = output(capsys)
    assert value["ok"] is False
    assert value["error"]["code"] == "validation_error"


def test_plan_app_must_match_command(make_docx, tmp_path, capsys):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    plan = write_plan(tmp_path / "plan.json", target, tmp_path / "final.docx")
    assert main(["slides", "apply", "--plan", str(plan), "--dry-run"], adapters=adapters()) == 2
    assert output(capsys)["error"]["details"]["plan_app"] == "word"
