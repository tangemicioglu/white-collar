from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from conftest import write_plan
from whitecollar.adapters.word import OoxmlWordAdapter
from whitecollar.engine import RuntimeAdapters, apply_plan, inspect_document
from whitecollar.errors import ValidationError
from whitecollar.models import Plan


class NeverUsed:
    pass


def runtime() -> RuntimeAdapters:
    return RuntimeAdapters(OoxmlWordAdapter(), NeverUsed(), NeverUsed())


def load_plan(path: Path) -> Plan:
    return Plan.from_dict(json.loads(path.read_text(encoding="utf-8")))


def test_inspect_real_docx(make_docx, tmp_path):
    target = make_docx(tmp_path / "brief.docx", "Quarterly Draft", "Owner: Ada")
    data = inspect_document("word", target, "read-only", runtime())
    assert data["paragraphs"] == 2
    assert data["text"] == "Quarterly DraftOwner: Ada"
    assert len(data["sha256"]) == 64


def test_dry_run_reports_matches_without_writing(make_docx, tmp_path):
    target = make_docx(tmp_path / "brief.docx", "Draft", "Draft")
    output = tmp_path / "final.docx"
    plan = load_plan(write_plan(tmp_path / "plan.json", target, output))
    result = apply_plan(plan, dry_run=True, adapters=runtime())
    assert result["changes"][0]["matches"] == 2
    assert result["written"] is False
    assert not output.exists()


def test_review_save_as_is_end_to_end(make_docx, tmp_path):
    target = make_docx(tmp_path / "brief.docx", "Draft report", "Draft status")
    output = tmp_path / "final.docx"
    plan = load_plan(write_plan(tmp_path / "plan.json", target, output))
    result = apply_plan(plan, dry_run=False, adapters=runtime())
    assert output.exists()
    assert result["changes"][0]["matches"] == 2
    assert OoxmlWordAdapter().inspect(output)["text"] == "Final reportFinal status"
    assert OoxmlWordAdapter().inspect(target)["text"] == "Draft reportDraft status"


def test_edit_in_place_creates_snapshot(make_docx, tmp_path):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    snapshot = tmp_path / "brief.before.docx"
    plan = load_plan(write_plan(tmp_path / "plan.json", target, snapshot, policy="edit", mode="in-place"))
    result = apply_plan(plan, dry_run=False, adapters=runtime())
    assert result["snapshot"] == str(snapshot)
    assert OoxmlWordAdapter().inspect(target)["text"] == "Final"
    assert OoxmlWordAdapter().inspect(snapshot)["text"] == "Draft"


def test_expected_hash_prevents_stale_write(make_docx, tmp_path):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    plan = load_plan(write_plan(tmp_path / "plan.json", target, tmp_path / "final.docx", expected_sha256="0" * 64))
    with pytest.raises(ValidationError, match="does not match"):
        apply_plan(plan, dry_run=False, adapters=runtime())


def test_existing_destination_is_not_overwritten(make_docx, tmp_path):
    target = make_docx(tmp_path / "brief.docx", "Draft")
    output = make_docx(tmp_path / "final.docx", "Keep me")
    before = hashlib.sha256(output.read_bytes()).hexdigest()
    plan = load_plan(write_plan(tmp_path / "plan.json", target, output))
    with pytest.raises(ValidationError, match="already exists"):
        apply_plan(plan, dry_run=False, adapters=runtime())
    assert hashlib.sha256(output.read_bytes()).hexdigest() == before


def test_live_apply_rejects_unmatched_operation(make_docx, tmp_path):
    target = make_docx(tmp_path / "brief.docx", "Already final")
    output = tmp_path / "final.docx"
    plan = load_plan(write_plan(tmp_path / "plan.json", target, output))
    with pytest.raises(ValidationError, match="matched no text"):
        apply_plan(plan, dry_run=False, adapters=runtime())
    assert not output.exists()
