"""Narrow Outlook Classic COM adapter.

Message metadata and explicitly requested bodies are exposed for reads. The
write surface is a finite semantic catalog: mark read/unread, move to a
standard folder, delete to Outlook's Deleted Items behavior, and create a
draft. It never dispatches arbitrary COM methods, replies, forwards, or sends
newly composed mail.
"""

from __future__ import annotations

import re
from typing import Any, Callable

from ..errors import BackendUnavailableError, ValidationError
from ..mail_ops import MAIL_COM_OPERATIONS


_DEFAULT_FOLDER_IDS = {
    "inbox": 6,
    "sent items": 5,
    "drafts": 16,
    "deleted items": 3,
    "outbox": 4,
    "junk email": 23,
}
_TOKEN_PATTERN = re.compile(r'"([^"]+)"|(\S+)')


class OutlookComAdapter:
    """Semantic adapter for the current user's Outlook Classic profile."""

    def __init__(self, *, app_factory: Callable[[], Any] | None = None, default_folder: str = "Inbox") -> None:
        self._app_factory = app_factory or _default_outlook_app
        self._default_folder = default_folder

    def search(self, query: str, *, limit: int, folder: str = "Inbox") -> list[dict[str, Any]]:
        app = self._get_app()
        namespace = _namespace(app)
        folder_object = _default_folder(namespace, folder or self._default_folder)
        terms = _parse_query(query)
        matches: list[dict[str, Any]] = []
        for item in _iter_collection(_safe_value(folder_object, "Items", [])):
            if not _is_mail_item(item) or not _matches(item, terms):
                continue
            matches.append(_message_metadata(item, folder=folder))
            if len(matches) >= limit:
                break
        return matches

    def read(self, message_id: str, *, include_body: bool = False) -> dict[str, Any]:
        app = self._get_app()
        item = _get_item_from_id(_namespace(app), message_id)
        if not _is_mail_item(item):
            raise ValidationError("Outlook item is not a mail message", details={"message_id": message_id})
        return _message_metadata(item, include_body=include_body)

    def apply(self, plan: Any, *, dry_run: bool) -> dict[str, Any]:
        message_id = getattr(getattr(plan, "target", None), "id", None)
        if not isinstance(message_id, str) or not message_id.strip():
            raise ValidationError("mail plans require a non-empty target.id")
        app = self._get_app()
        operations = getattr(plan, "operations", ())
        if any(operation["op"] == "mail_live_create_draft" for operation in operations):
            if len(operations) != 1:
                raise ValidationError("mail_live_create_draft must be a standalone plan")
            return self._create_draft(app, message_id, operations[0].get("args", {}), dry_run=dry_run)
        namespace = _namespace(app)
        item = _get_item_from_id(namespace, message_id)
        if not _is_mail_item(item):
            raise ValidationError("Outlook item is not a mail message", details={"message_id": message_id})
        changes: list[dict[str, Any]] = []
        for operation in getattr(plan, "operations", ()):
            name = operation["op"]
            if name not in MAIL_COM_OPERATIONS:
                raise ValidationError("unsupported Outlook mail operation", details={"operation": name})
            if name in {"mail_live_mark_read", "mail_live_mark_unread"}:
                before = bool(_safe_value(item, "UnRead", False))
                after = name == "mail_live_mark_unread"
                if not dry_run:
                    item.UnRead = after
                    save = getattr(item, "Save", None)
                    if callable(save):
                        save()
                changes.append({"op": name, "id": message_id, "unread_before": before, "unread_after": after})
                continue
            if name == "mail_live_move":
                folder = operation.get("args", {}).get("folder")
                destination = _default_folder(namespace, folder)
                if not dry_run:
                    move = getattr(item, "Move", None)
                    if not callable(move):
                        raise ValidationError("Outlook message does not support Move", details={"message_id": message_id})
                    moved = move(destination)
                    message_id = str(_safe_value(moved, "EntryID", message_id))
                changes.append({"op": name, "id": message_id, "folder": folder})
                continue
            if name == "mail_live_delete":
                parent = _safe_value(item, "Parent", None)
                if str(_safe_value(parent, "Name", "")).strip().lower() == "deleted items":
                    raise ValidationError(
                        "refusing to delete a message already in Deleted Items",
                        details={"message_id": message_id},
                    )
                if not dry_run:
                    delete = getattr(item, "Delete", None)
                    if not callable(delete):
                        raise ValidationError("Outlook message does not support Delete", details={"message_id": message_id})
                    delete()
                changes.append({"op": name, "id": message_id, "deleted": not dry_run})
                continue
            if name == "mail_live_send":
                if bool(_safe_value(item, "Sent", False)):
                    raise ValidationError(
                        "refusing to send a message that Outlook already marks as sent",
                        details={"message_id": message_id},
                    )
                if not dry_run:
                    send = getattr(item, "Send", None)
                    if not callable(send):
                        raise ValidationError("Outlook message does not support Send", details={"message_id": message_id})
                    send()
                changes.append({"op": name, "id": message_id, "sent": not dry_run})
        return {"changes": changes, "written": not dry_run}

    def _create_draft(self, app: Any, mailbox: str, args: dict[str, Any], *, dry_run: bool) -> dict[str, Any]:
        account = mailbox.strip()
        if account.lower() != "mailbox":
            _find_account(app, account)
        summary: dict[str, Any] = {
            "op": "mail_live_create_draft",
            "mailbox": account,
            "to": args["to"],
            "subject": args["subject"],
            "body_chars": len(args["body"]),
            "created": False,
        }
        for field in ("cc", "bcc"):
            if args.get(field):
                summary[field] = args[field]
        if dry_run:
            return {"changes": [summary], "written": False}
        create_item = getattr(app, "CreateItem", None)
        if not callable(create_item):
            raise ValidationError("Outlook application does not support draft creation")
        try:
            draft = create_item(0)  # Outlook MailItem
            if account.lower() != "mailbox":
                draft.SendUsingAccount = _find_account(app, account)
            draft.To = args["to"]
            draft.Subject = args["subject"]
            draft.Body = args["body"]
            if "cc" in args:
                draft.CC = args["cc"]
            if "bcc" in args:
                draft.BCC = args["bcc"]
            save = getattr(draft, "Save", None)
            if not callable(save):
                raise ValidationError("Outlook draft does not support Save")
            save()
        except ValidationError:
            raise
        except Exception as exc:
            raise BackendUnavailableError("outlook-com") from exc
        draft_id = str(_safe_value(draft, "EntryID", ""))
        if not draft_id:
            raise ValidationError("Outlook did not return an entry ID for the new draft")
        summary["draft_id"] = draft_id
        summary["created"] = True
        return {"changes": [summary], "written": True}

    def _get_app(self) -> Any:
        try:
            return self._app_factory()
        except ImportError as exc:
            raise BackendUnavailableError("outlook-com") from exc
        except OSError as exc:
            raise BackendUnavailableError("outlook-com") from exc


def _default_outlook_app() -> Any:
    try:
        from win32com.client import Dispatch, GetActiveObject  # type: ignore[import-not-found]
    except ImportError as exc:
        raise BackendUnavailableError("outlook-com") from exc
    try:
        return GetActiveObject("Outlook.Application")
    except Exception as exc:
        try:
            return Dispatch("Outlook.Application")
        except Exception as dispatch_exc:
            raise BackendUnavailableError("outlook-com") from dispatch_exc


def _namespace(app: Any) -> Any:
    getter = getattr(app, "GetNamespace", None)
    if not callable(getter):
        raise ValidationError("Outlook application does not expose the MAPI namespace")
    try:
        return getter("MAPI")
    except Exception as exc:
        raise BackendUnavailableError("outlook-com") from exc


def _default_folder(namespace: Any, folder: str) -> Any:
    key = folder.strip().lower()
    folder_id = _DEFAULT_FOLDER_IDS.get(key)
    if folder_id is None:
        raise ValidationError(
            "unsupported Outlook folder",
            details={"folder": folder, "supported": sorted(_DEFAULT_FOLDER_IDS)},
        )
    try:
        return namespace.GetDefaultFolder(folder_id)
    except Exception as exc:
        raise BackendUnavailableError("outlook-com") from exc


def _get_item_from_id(namespace: Any, message_id: str) -> Any:
    try:
        item = namespace.GetItemFromID(message_id)
    except Exception as exc:
        raise ValidationError("Outlook message was not found", details={"message_id": message_id}) from exc
    if item is None:
        raise ValidationError("Outlook message was not found", details={"message_id": message_id})
    return item


def _find_account(app: Any, address: str) -> Any:
    session = _safe_value(app, "Session", None)
    accounts = _safe_value(session, "Accounts", None)
    for account in _iter_collection(accounts):
        smtp = str(_safe_value(account, "SmtpAddress", "")).strip().lower()
        name = str(_safe_value(account, "AccountName", "")).strip().lower()
        if address.lower() in {smtp, name}:
            return account
    raise ValidationError(
        "Outlook account was not found",
        details={"account": address},
    )


def _message_metadata(item: Any, *, folder: str | None = None, include_body: bool = False) -> dict[str, Any]:
    attachments = _safe_value(_safe_value(item, "Attachments", None), "Count", 0)
    value: dict[str, Any] = {
        "id": str(_safe_value(item, "EntryID", "")),
        "subject": str(_safe_value(item, "Subject", "")),
        "sender": {
            "name": str(_safe_value(item, "SenderName", "")),
            "address": str(_safe_value(item, "SenderEmailAddress", "")),
        },
        "to": str(_safe_value(item, "To", "")),
        "cc": str(_safe_value(item, "CC", "")),
        "received": _iso_value(_safe_value(item, "ReceivedTime", None)),
        "sent": _iso_value(_safe_value(item, "SentOn", None)),
        "unread": bool(_safe_value(item, "UnRead", False)),
        "size_bytes": int(_safe_value(item, "Size", 0) or 0),
        "has_attachments": int(attachments or 0) > 0,
        "attachment_count": int(attachments or 0),
    }
    if folder is not None:
        value["folder"] = folder
    if include_body:
        value["body"] = str(_safe_value(item, "Body", ""))
    return value


def _parse_query(query: str) -> dict[str, list[str] | bool]:
    terms: dict[str, list[str] | bool] = {"text": [], "from": [], "to": [], "cc": [], "subject": [], "attachment": False}
    for match in _TOKEN_PATTERN.finditer(query):
        token = (match.group(1) or match.group(2) or "").strip()
        lowered = token.lower()
        if lowered in {"has:attachment", "has:attachments"}:
            terms["attachment"] = True
        elif ":" in lowered:
            field, value = lowered.split(":", 1)
            if field in {"from", "to", "cc", "subject"} and value:
                terms[field].append(value)  # type: ignore[union-attr]
            else:
                terms["text"].append(lowered)  # type: ignore[union-attr]
        else:
            terms["text"].append(lowered)  # type: ignore[union-attr]
    return terms


def _matches(item: Any, terms: dict[str, list[str] | bool]) -> bool:
    sender = " ".join((str(_safe_value(item, "SenderName", "")), str(_safe_value(item, "SenderEmailAddress", "")))).lower()
    subject = str(_safe_value(item, "Subject", "")).lower()
    recipient = " ".join((str(_safe_value(item, "To", "")), str(_safe_value(item, "CC", "")))).lower()
    attachment_count = int(_safe_value(_safe_value(item, "Attachments", None), "Count", 0) or 0)
    if terms["attachment"] and attachment_count == 0:
        return False
    for value in terms["from"]:  # type: ignore[union-attr]
        if value not in sender:
            return False
    for value in terms["to"]:  # type: ignore[union-attr]
        if value not in recipient:
            return False
    for value in terms["cc"]:  # type: ignore[union-attr]
        if value not in recipient:
            return False
    for value in terms["subject"]:  # type: ignore[union-attr]
        if value not in subject:
            return False
    searchable = " ".join((subject, sender, recipient))
    return all(value in searchable for value in terms["text"])  # type: ignore[union-attr]


def _is_mail_item(item: Any) -> bool:
    message_class = str(_safe_value(item, "MessageClass", "IPM.Note"))
    return message_class.startswith("IPM.Note") and bool(_safe_value(item, "EntryID", ""))


def _iter_collection(collection: Any):
    if collection is None:
        return iter(())
    try:
        return iter(collection)
    except TypeError:
        count = int(_safe_value(collection, "Count", 0) or 0)
        item = getattr(collection, "Item", None)
        if not callable(item):
            item = collection if callable(collection) else None
        if item is None:
            return iter(())
        return (item(index) for index in range(1, count + 1))


def _safe_value(value: Any, attribute: str, default: Any = None) -> Any:
    if value is None:
        return default
    try:
        return getattr(value, attribute)
    except Exception:
        return default


def _iso_value(value: Any) -> str | None:
    if value is None:
        return None
    isoformat = getattr(value, "isoformat", None)
    if callable(isoformat):
        return str(isoformat())
    return str(value)
