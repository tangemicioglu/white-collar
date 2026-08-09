from __future__ import annotations

from pathlib import Path

import pytest

from whitecollar.errors import PolicyError
from whitecollar.permissions import (
    CAPABILITIES,
    OPERATION_CAPABILITIES,
    PROFILE_CAPABILITIES,
    capability_for_operation,
    catalog,
    decide,
    require_capability,
    setup_capabilities,
    SETUP_PRESETS,
)


def test_every_registered_office_operation_has_a_capability():
    assert CAPABILITIES
    assert all(
        capability_for_operation(app, operation) in CAPABILITIES
        for (app, operation) in OPERATION_CAPABILITIES
    )


def test_profiles_are_monotonic_for_office_capabilities():
    assert PROFILE_CAPABILITIES["read-only"] < PROFILE_CAPABILITIES["review"]
    assert PROFILE_CAPABILITIES["review"] < PROFILE_CAPABILITIES["edit"]
    assert "mail.write.state" in PROFILE_CAPABILITIES["review"]
    assert "mail.write.state" not in PROFILE_CAPABILITIES["read-only"]
    assert "mail.write.organize" not in PROFILE_CAPABILITIES["review"]
    assert "mail.write.organize" in PROFILE_CAPABILITIES["edit"]
    assert "mail.write.compose" not in PROFILE_CAPABILITIES["review"]
    assert "mail.write.compose" in PROFILE_CAPABILITIES["edit"]
    assert "mail.write.send" not in PROFILE_CAPABILITIES["edit"]
    assert "mail.write.send" in PROFILE_CAPABILITIES["send"]


def test_read_only_allows_targeted_word_reads(tmp_path: Path):
    target = str((tmp_path / "brief.docx").resolve())
    decision = decide("read-only", "word.read", target=target)
    assert decision["ok"] is True
    assert decision["decision"] == "allow"


def test_file_capability_requires_absolute_target():
    decision = decide("review", "word.write.save_as", target="relative.docx")
    assert decision["ok"] is False
    assert "absolute" in decision["reason"]


def test_sensitive_mail_body_is_not_read_only():
    denied = decide("read-only", "mail.body.read", target="message-1")
    assert denied["ok"] is False
    assert denied["decision"] == "deny"
    assert decide("review", "mail.body.read", target="message-1")["ok"] is True


def test_unknown_capability_fails_closed():
    with pytest.raises(PolicyError, match="unknown capability"):
        require_capability("review", "mail.raw.com.method", target="message-1")


def test_catalog_is_versioned_and_exposes_grants():
    value = catalog(policy="read-only")
    assert value["schema"] == "white-collar.permissions/v1"
    mail_body = next(item for item in value["capabilities"] if item["name"] == "mail.body.read")
    assert mail_body["granted"] is False


def test_setup_profiles_are_bounded_per_application():
    assert setup_capabilities("mail", "review") == (
        "mail.body.read",
        "mail.metadata.read",
        "mail.write.state",
    )
    assert "mail.write.compose" in setup_capabilities("mail", "edit")
    assert "mail.write.send" in setup_capabilities("mail", "send")
    assert setup_capabilities("mail", "disabled") == ()
    with pytest.raises(PolicyError, match="not available"):
        setup_capabilities("word", "send")


def test_setup_presets_are_explicit_and_mail_send_is_opt_in():
    assert SETUP_PRESETS["safe"] == {"word": "review", "slides": "review", "mail": "disabled"}
    assert SETUP_PRESETS["office-authoring"]["mail"] == "disabled"
    assert SETUP_PRESETS["outlook-review"] == {"mail": "review"}
    assert SETUP_PRESETS["outlook-send"] == {"mail": "send"}
