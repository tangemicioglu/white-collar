from __future__ import annotations

from whitecollar.authority import Authority, Grant
from whitecollar.doctor import diagnose


def test_diagnose_reports_permission_counts_without_targets():
    authority = Authority(
        (
            Grant("mail", "com", "edit", ("mail.write.compose",), ("mailbox",)),
        ),
        source="test",
    )
    value = diagnose(authority, module_available=lambda name: name == "PIL")
    assert value["environment"]["pillow"] is True
    assert value["permissions"]["owner_grants"] == 1
    assert value["permissions"]["owner_apps"] == ["mail"]
    assert value["permissions"]["owner_policies"] == ["edit"]
    assert value["permissions"]["targets"].startswith("redacted")
    assert value["safety"]["office_probe"] == "not-run"
