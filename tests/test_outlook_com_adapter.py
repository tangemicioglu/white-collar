from __future__ import annotations

import datetime as dt
from types import SimpleNamespace

import pytest

from whitecollar.adapters.outlook import OutlookComAdapter
from whitecollar.errors import ValidationError


class FakeMessage:
    MessageClass = "IPM.Note"

    def __init__(self, message_id: str, subject: str, sender: str, address: str, body: str, *, attachments: int = 0):
        self.EntryID = message_id
        self.Subject = subject
        self.SenderName = sender
        self.SenderEmailAddress = address
        self.To = "team@example.com"
        self.CC = ""
        self.ReceivedTime = dt.datetime(2026, 8, 9, 8, 30)
        self.SentOn = dt.datetime(2026, 8, 9, 8, 0)
        self.UnRead = True
        self.Size = 2048
        self.Attachments = SimpleNamespace(Count=attachments)
        self._body = body
        self.body_reads = 0

    @property
    def Body(self):
        self.body_reads += 1
        return self._body


class FakeFolder:
    def __init__(self, *items):
        self.Name = "Inbox"
        self.Items = list(items)


class FakeNamespace:
    def __init__(self, folder, items):
        self.folder = folder
        self.items = {item.EntryID: item for item in items}

    def GetDefaultFolder(self, folder_id):
        assert folder_id == 6
        return self.folder

    def GetItemFromID(self, message_id):
        if message_id not in self.items:
            raise KeyError(message_id)
        return self.items[message_id]


class FakeOutlook:
    def __init__(self, namespace):
        self.namespace = namespace

    def GetNamespace(self, name):
        assert name == "MAPI"
        return self.namespace


def adapter_and_messages():
    first = FakeMessage("m-1", "Roadmap review", "Ada Lovelace", "ada@example.com", "Secret body", attachments=1)
    second = FakeMessage("m-2", "Unrelated", "Grace Hopper", "grace@example.com", "Other body")
    folder = FakeFolder(first, second)
    namespace = FakeNamespace(folder, [first, second])
    return OutlookComAdapter(app_factory=lambda: FakeOutlook(namespace)), first, second


def test_search_is_metadata_only_and_supports_narrow_query_fields():
    adapter, first, second = adapter_and_messages()
    value = adapter.search("from:ada@example.com roadmap", limit=10)
    assert [item["id"] for item in value] == ["m-1"]
    assert value[0]["sender"]["address"] == "ada@example.com"
    assert value[0]["has_attachments"] is True
    assert "body" not in value[0]
    assert first.body_reads == 0
    assert second.body_reads == 0


def test_read_body_is_accessed_only_when_explicitly_requested():
    adapter, first, _ = adapter_and_messages()
    metadata = adapter.read("m-1")
    assert "body" not in metadata
    assert first.body_reads == 0
    full = adapter.read("m-1", include_body=True)
    assert full["body"] == "Secret body"
    assert first.body_reads == 1


def test_read_missing_message_and_unknown_folder_are_validation_errors():
    adapter, _, _ = adapter_and_messages()
    with pytest.raises(ValidationError, match="not found"):
        adapter.read("missing")
    with pytest.raises(ValidationError, match="unsupported Outlook folder"):
        adapter.search("roadmap", limit=10, folder="Archive")
