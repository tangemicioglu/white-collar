from __future__ import annotations

import json
from pathlib import Path

import pytest

from whitecollar.errors import PolicyError, ValidationError
from whitecollar.models import PLAN_SCHEMA, Plan
from whitecollar.policy import authorize_plan


def test_fixture_plan_is_valid_and_versioned():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    plan = Plan.from_dict(raw)
    assert plan.schema == PLAN_SCHEMA
    assert plan.target.path == r"C:\fixtures\input.docx"
    assert plan.operations[0]["occurrence"] == "all"


def test_unknown_plan_fields_are_rejected():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["com_method"] = "Anything"
    with pytest.raises(ValidationError, match="unknown field"):
        Plan.from_dict(raw)


def test_read_only_denies_even_a_dry_run():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["policy"] = "read-only"
    plan = Plan.from_dict(raw)
    with pytest.raises(PolicyError, match="including dry-runs"):
        authorize_plan(plan, dry_run=True)


def test_review_allows_save_as_but_denies_in_place():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    plan = Plan.from_dict(raw)
    authorize_plan(plan, dry_run=False)
    raw["write"] = {"mode": "in-place", "snapshot": "C:/fixtures/snapshot.docx"}
    in_place = Plan.from_dict(raw)
    with pytest.raises(PolicyError, match="in-place"):
        authorize_plan(in_place, dry_run=False)


def test_edit_allows_in_place_only_with_explicit_snapshot():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["policy"] = "edit"
    raw["write"] = {"mode": "in-place", "snapshot": "C:/fixtures/snapshot.docx"}
    authorize_plan(Plan.from_dict(raw), dry_run=False)
    del raw["write"]["snapshot"]
    with pytest.raises(ValidationError, match="non-empty string"):
        Plan.from_dict(raw)


def test_word_com_read_operation_uses_read_only_and_no_write_intent():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["operations"] = [{"op": "word_live_get_text", "args": {}}]
    raw["policy"] = "read-only"
    raw["write"] = {"mode": "none"}
    authorize_plan(Plan.from_dict(raw), dry_run=False)


def test_powerpoint_com_read_operation_uses_read_only_and_no_write_intent():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["app"] = "slides"
    raw["operations"] = [{"op": "slides_live_get_info", "args": {}}]
    raw["policy"] = "read-only"
    raw["write"] = {"mode": "none"}
    authorize_plan(Plan.from_dict(raw), dry_run=False)


def test_powerpoint_com_mutation_requires_explicit_write_policy():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["app"] = "slides"
    raw["operations"] = [{"op": "slides_live_set_title", "args": {"title": "Final"}}]
    raw["policy"] = "review"
    raw["write"] = {"mode": "save-as", "path": "C:/fixtures/review.pptx"}
    authorize_plan(Plan.from_dict(raw), dry_run=False)
    raw["write"] = {"mode": "in-place", "snapshot": "C:/fixtures/before.pptx"}
    with pytest.raises(PolicyError, match="in-place"):
        authorize_plan(Plan.from_dict(raw), dry_run=False)


def test_schema_documents_are_valid_json():
    for path in Path("schemas").glob("*.schema.json"):
        assert json.loads(path.read_text(encoding="utf-8"))["$schema"].endswith("2020-12/schema")
