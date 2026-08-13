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


def test_remove_watermark_is_a_bounded_word_mutation():
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["operations"] = [{"op": "word_live_remove_watermark", "args": {"text": "DRAFT"}}]
    authorize_plan(Plan.from_dict(raw), dry_run=False)
    raw["policy"] = "read-only"
    with pytest.raises(PolicyError, match="including dry-runs"):
        authorize_plan(Plan.from_dict(raw), dry_run=True)


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


@pytest.mark.parametrize(
    "operation,args,match",
    [
        ("word_live_apply_style", {"style_name": "Title", "unexpected": True}, "unknown field"),
        ("word_live_remove_hyperlink", {}, "identify a hyperlink"),
        ("word_live_set_protection", {"protection_type": "admin"}, "invalid"),
        ("word_live_compare_documents", {"source_path": "relative.docx"}, "absolute path"),
    ],
)
def test_new_word_operations_have_bounded_argument_validation(operation, args, match):
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["operations"] = [{"op": operation, "args": args}]
    with pytest.raises(ValidationError, match=match):
        Plan.from_dict(raw)


@pytest.mark.parametrize(
    "operation,args,match",
    [
        ("slides_live_group", {"shape_names": []}, "non-empty"),
        ("slides_live_ungroup", {}, "identify a shape"),
        ("slides_live_crop_image", {"shape_name": "Picture"}, "crop value"),
        ("slides_live_add_media", {"media_path": "relative.wav"}, "absolute path"),
    ],
)
def test_new_slides_operations_have_bounded_argument_validation(operation, args, match):
    raw = json.loads(Path("tests/fixtures/word-replace-plan.json").read_text(encoding="utf-8"))
    raw["app"] = "slides"
    raw["operations"] = [{"op": operation, "args": args}]
    with pytest.raises(ValidationError, match=match):
        Plan.from_dict(raw)


def test_mail_write_plan_splits_review_state_from_edit_operations_and_has_no_file_write_intent():
    raw = {
        "schema": "white-collar.plan/v1",
        "app": "mail",
        "target": {"id": "message-1"},
        "policy": "review",
        "operations": [{"op": "mail_live_mark_read"}],
        "write": {"mode": "none"},
    }
    plan = Plan.from_dict(raw)
    authorize_plan(plan, dry_run=True)
    raw["operations"] = [{"op": "mail_live_move", "args": {"folder": "Sent Items"}}]
    with pytest.raises(PolicyError, match="capability"):
        authorize_plan(Plan.from_dict(raw), dry_run=False)
    raw["policy"] = "edit"
    authorize_plan(Plan.from_dict(raw), dry_run=True)
    raw["write"] = {"mode": "save-as", "path": "C:/not-a-mail-file.pst"}
    with pytest.raises(ValidationError, match="mail plans"):
        Plan.from_dict(raw)


def test_mail_composition_is_edit_only_and_requires_explicit_fields():
    raw = {
        "schema": "white-collar.plan/v1",
        "app": "mail",
        "target": {"id": "mailbox"},
        "policy": "review",
        "operations": [{
            "op": "mail_live_create_draft",
            "args": {"to": "person@example.com", "subject": "Test", "body": "Hello"},
        }],
        "write": {"mode": "none"},
    }
    with pytest.raises(PolicyError, match="'edit' policy"):
        authorize_plan(Plan.from_dict(raw), dry_run=True)
    raw["policy"] = "edit"
    authorize_plan(Plan.from_dict(raw), dry_run=True)
    raw["operations"][0]["args"]["unexpected"] = "nope"
    with pytest.raises(ValidationError, match="unknown field"):
        Plan.from_dict(raw)


def test_mail_send_plan_requires_the_send_policy_level():
    raw = {
        "schema": "white-collar.plan/v1",
        "app": "mail",
        "target": {"id": "draft-1"},
        "policy": "edit",
        "operations": [{"op": "mail_live_send"}],
        "write": {"mode": "none"},
    }
    with pytest.raises(PolicyError, match="'send' policy"):
        authorize_plan(Plan.from_dict(raw), dry_run=True)
    raw["policy"] = "send"
    authorize_plan(Plan.from_dict(raw), dry_run=True)


def test_schema_documents_are_valid_json():
    for path in Path("schemas").glob("*.schema.json"):
        assert json.loads(path.read_text(encoding="utf-8"))["$schema"].endswith("2020-12/schema")
