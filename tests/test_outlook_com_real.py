from __future__ import annotations

import pytest

from whitecollar.adapters.outlook import OutlookComAdapter
from whitecollar.errors import BackendUnavailableError


@pytest.mark.real_outlook
def test_outlook_metadata_search_against_current_profile():
    """Connect to Outlook without reading bodies or mutating the mailbox."""

    try:
        results = OutlookComAdapter().search(
            "subject:white-collar-no-such-message-9f5f0ec4",
            limit=1,
            folder="Inbox",
        )
    except BackendUnavailableError as exc:
        pytest.skip(str(exc))
    assert isinstance(results, list)
