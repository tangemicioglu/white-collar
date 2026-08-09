from __future__ import annotations

import json

import pytest

from whitecollar.authority import (
    GRANT_SCHEMA,
    Authority,
    MemoryCredentialStore,
    load_authority,
    make_grant,
    revoke_grant,
    save_grant,
)
from whitecollar.errors import PolicyError, ValidationError


def test_default_authority_allows_review_writes_but_disables_edit_and_outlook(tmp_path):
    target = str((tmp_path / "brief.docx").resolve())
    authority = Authority.default()
    authority.require_policy("word", "read-only", target=target)
    authority.require_backend("word", "com")
    authority.require_policy("word", "review", target=target)
    with pytest.raises(PolicyError, match="not been approved"):
        authority.require_policy("word", "edit", target=target)
    with pytest.raises(PolicyError, match="not been approved"):
        authority.require_backend("mail", "com")


def test_owner_grant_is_loaded_from_protected_store_and_is_target_scoped(tmp_path):
    target = str((tmp_path / "brief.docx").resolve())
    other = str((tmp_path / "other.docx").resolve())
    store = MemoryCredentialStore()
    grant = make_grant(
        app="word",
        backend="local",
        policy="review",
        targets=[target],
        capabilities=["word.capture"],
    )
    save_grant(Authority.default(), grant, store)
    authority = load_authority(store=store)
    authority.require_access("word", "local", "review", target, ("word.capture",))
    with pytest.raises(PolicyError, match="not been approved"):
        authority.require_access("word", "local", "review", other, ("word.capture",))


def test_authority_files_are_rejected_instead_of_being_a_grant(tmp_path):
    path = tmp_path / "authority.json"
    path.write_text("{}", encoding="utf-8")
    with pytest.raises(ValidationError, match="authority files are not supported"):
        load_authority(path)


def test_protected_payload_is_versioned_and_rejects_profile_escalation():
    raw = {"schema": GRANT_SCHEMA, "grants": []}
    assert Authority.from_grant_dict(json.loads(json.dumps(raw)), source="test").owner_grants == ()
    with pytest.raises(ValidationError, match="exceeds its policy"):
        Authority.from_grant_dict(
            {
                "schema": GRANT_SCHEMA,
                "grants": [
                    {
                        "app": "word",
                        "backend": "local",
                        "policy": "read-only",
                        "capabilities": ["word.write.save_as"],
                        "targets": ["*"],
                    }
                ],
            },
            source="test",
        )


def test_revoke_can_remove_one_capability_from_a_broad_owner_grant(tmp_path):
    target = str((tmp_path / "brief.docx").resolve())
    store = MemoryCredentialStore()
    broad = make_grant(app="word", backend="local", policy="review", targets=[target])
    authority = save_grant(Authority.default(), broad, store)
    narrow = make_grant(
        app="word",
        backend="local",
        policy="review",
        targets=[target],
        capabilities=["word.capture"],
    )
    updated = revoke_grant(authority, narrow, store)
    updated.require_access("word", "local", "review", target, ("word.read",))
    with pytest.raises(PolicyError, match="not been approved"):
        updated.require_access("word", "local", "review", target, ("word.capture",))


def test_send_grant_is_a_distinct_exact_target_level():
    store = MemoryCredentialStore()
    authority = save_grant(
        Authority.default(),
        make_grant(app="mail", backend="com", policy="send", targets=["draft-1"]),
        store,
    )
    authority.require_access("mail", "com", "send", "draft-1", ("mail.write.send",))
    with pytest.raises(PolicyError, match="not been approved"):
        authority.require_access("mail", "com", "send", "draft-2", ("mail.write.send",))
