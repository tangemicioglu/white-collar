from __future__ import annotations

import datetime as dt
from types import SimpleNamespace

import pytest

from whitecollar.adapters.outlook import OutlookComAdapter
from whitecollar.errors import ValidationError
from whitecollar.models import Plan


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
        self.save_calls = 0
        self.deleted = False
        self.Sent = False
        self.send_calls = 0
        self.Parent = None

    @property
    def Body(self):
        self.body_reads += 1
        return self._body

    def Save(self):
        self.save_calls += 1

    def Move(self, folder):
        self.Parent = folder
        return self

    def Delete(self):
        self.deleted = True

    def Send(self):
        self.send_calls += 1
        self.Sent = True


class FakeDraft:
    MessageClass = "IPM.Note"

    def __init__(self, message_id="draft-1"):
        self.EntryID = message_id
        self.To = ""
        self.CC = ""
        self.BCC = ""
        self.Subject = ""
        self.Body = ""
        self.SendUsingAccount = None
        self.save_calls = 0

    def Save(self):
        self.save_calls += 1


class FakeFolder:
    def __init__(self, *items, name="Inbox"):
        self.Name = name
        self.Items = list(items)
        for item in self.Items:
            item.Parent = self


class FakeNamespace:
    def __init__(self, folder, items):
        self.folder = folder
        self.items = {item.EntryID: item for item in items}
        self.folders = {
            5: FakeFolder(name="Sent Items"),
            6: folder,
        }

    def GetDefaultFolder(self, folder_id):
        return self.folders[folder_id]

    def GetItemFromID(self, message_id):
        if message_id not in self.items:
            raise KeyError(message_id)
        return self.items[message_id]


class FakeOutlook:
    def __init__(self, namespace):
        self.namespace = namespace
        self.created_drafts = []
        self.Session = SimpleNamespace(
            Accounts=[SimpleNamespace(SmtpAddress="tgemicioglu@outlook.com", AccountName="tgemicioglu@outlook.com")]
        )

    def GetNamespace(self, name):
        assert name == "MAPI"
        return self.namespace

    def CreateItem(self, item_type):
        assert item_type == 0
        draft = FakeDraft()
        self.created_drafts.append(draft)
        return draft


def adapter_and_messages():
    first = FakeMessage("m-1", "Roadmap review", "Ada Lovelace", "ada@example.com", "Secret body", attachments=1)
    second = FakeMessage("m-2", "Unrelated", "Grace Hopper", "grace@example.com", "Other body")
    folder = FakeFolder(first, second)
    namespace = FakeNamespace(folder, [first, second])
    return OutlookComAdapter(app_factory=lambda: FakeOutlook(namespace)), first, second


def mail_plan(
    operation: str,
    *,
    args: dict | None = None,
    policy: str = "edit",
    target_id: str = "m-1",
) -> Plan:
    return Plan.from_dict(
        {
            "schema": "white-collar.plan/v1",
            "app": "mail",
            "target": {"id": target_id},
            "policy": policy,
            "operations": [{"op": operation, "args": args or {}}],
            "write": {"mode": "none"},
        }
    )


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


def test_write_operations_are_semantic_and_dry_run_is_non_mutating():
    adapter, first, _ = adapter_and_messages()
    value = adapter.apply(mail_plan("mail_live_mark_read"), dry_run=True)
    assert value["written"] is False
    assert value["changes"][0]["unread_after"] is False
    assert first.UnRead is True
    assert first.save_calls == 0


def test_write_operations_mark_move_and_delete_real_fake_items():
    adapter, first, _ = adapter_and_messages()
    adapter.apply(mail_plan("mail_live_mark_read"), dry_run=False)
    assert first.UnRead is False
    assert first.save_calls == 1
    adapter.apply(mail_plan("mail_live_mark_unread"), dry_run=False)
    assert first.UnRead is True
    adapter.apply(mail_plan("mail_live_move", args={"folder": "Sent Items"}), dry_run=False)
    assert first.Parent.Name == "Sent Items"
    adapter.apply(mail_plan("mail_live_delete"), dry_run=False)
    assert first.deleted is True


def test_send_requires_send_plan_level_and_dry_run_does_not_send():
    adapter, first, _ = adapter_and_messages()
    value = adapter.apply(mail_plan("mail_live_send", policy="send"), dry_run=True)
    assert value["changes"][0]["sent"] is False
    assert first.send_calls == 0
    adapter.apply(mail_plan("mail_live_send", policy="send"), dry_run=False)
    assert first.send_calls == 1
    with pytest.raises(ValidationError, match="already marks as sent"):
        adapter.apply(mail_plan("mail_live_send", policy="send"), dry_run=False)


def test_create_draft_is_bounded_and_dry_run_does_not_create():
    first = FakeMessage("m-1", "Roadmap review", "Ada Lovelace", "ada@example.com", "Secret body")
    namespace = FakeNamespace(FakeFolder(first), [first])
    outlook = FakeOutlook(namespace)
    adapter = OutlookComAdapter(app_factory=lambda: outlook)
    plan = mail_plan(
        "mail_live_create_draft",
        policy="edit",
        target_id="tgemicioglu@outlook.com",
        args={
            "to": "person@example.com",
            "cc": "copy@example.com",
            "bcc": "blind@example.com",
            "subject": "A bounded test",
            "body": "Hello from a draft.",
        },
    )
    preview = adapter.apply(plan, dry_run=True)
    assert preview["written"] is False
    assert preview["changes"][0]["created"] is False
    assert not outlook.created_drafts

    result = adapter.apply(plan, dry_run=False)
    assert result["written"] is True
    draft = outlook.created_drafts[0]
    assert draft.To == "person@example.com"
    assert draft.CC == "copy@example.com"
    assert draft.BCC == "blind@example.com"
    assert draft.Subject == "A bounded test"
    assert draft.Body == "Hello from a draft."
    assert draft.SendUsingAccount.SmtpAddress == "tgemicioglu@outlook.com"
    assert draft.save_calls == 1
    assert result["changes"][0]["draft_id"] == "draft-1"


def test_create_draft_rejects_unknown_account():
    first = FakeMessage("m-1", "Roadmap review", "Ada Lovelace", "ada@example.com", "Secret body")
    namespace = FakeNamespace(FakeFolder(first), [first])
    outlook = FakeOutlook(namespace)
    adapter = OutlookComAdapter(app_factory=lambda: outlook)
    plan = mail_plan(
        "mail_live_create_draft",
        target_id="missing@example.com",
        args={"to": "person@example.com", "subject": "Test", "body": "Hello"},
    )
    with pytest.raises(ValidationError, match="account was not found"):
        adapter.apply(plan, dry_run=True)
