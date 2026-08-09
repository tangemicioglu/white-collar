from __future__ import annotations

import json

import pytest

from whitecollar.authority import AUTHORITY_SCHEMA, Authority, load_authority
from whitecollar.errors import PolicyError


def test_default_authority_is_read_only_and_disables_outlook():
    authority = Authority.default()
    authority.require_policy("word", "read-only")
    authority.require_backend("word", "com")
    with pytest.raises(PolicyError, match="exceeds owner authority"):
        authority.require_policy("word", "review")
    with pytest.raises(PolicyError, match="disabled"):
        authority.require_backend("mail", "com")


def test_authority_file_can_grant_only_explicit_owner_values(tmp_path):
    path = tmp_path / "authority.json"
    path.write_text(
        json.dumps(
            {
                "schema": AUTHORITY_SCHEMA,
                "policies": {"word": "review", "mail": "review"},
                "backends": ["word:com", "mail:com"],
            }
        ),
        encoding="utf-8",
    )
    authority = load_authority(path)
    authority.require_policy("word", "review")
    authority.require_policy("mail", "review")
    authority.require_backend("mail", "com")
    with pytest.raises(PolicyError, match="exceeds owner authority"):
        authority.require_policy("word", "edit")
