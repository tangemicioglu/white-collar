# Outlook Classic adapter

The opt-in Outlook backend uses the current user's Outlook Classic MAPI profile
through COM. Its finite write catalog currently supports marking a message
read/unread, creating a draft, moving it to a standard folder, and deleting it
through Outlook's normal Deleted Items behavior. It does not reply, forward, or
download attachments, and it never exposes arbitrary COM dispatch. Sending is
a separate `send` permission level and is limited to an existing draft.

Outlook COM is disabled by default by the authority layer. An owner must
explicitly grant the exact Outlook capability and target through the human-only
`permissions grant` flow; selecting `--backend com` alone cannot enable it.

## Commands

```powershell
white-collar mail search --backend com --folder Inbox --query "from:ada@example.com roadmap"
white-collar mail search --backend com --folder "Sent Items" --query 'subject:"quarterly review"'
white-collar mail read --backend com --id MESSAGE_ENTRY_ID
white-collar mail read --backend com --id MESSAGE_ENTRY_ID --include-body --policy review
white-collar mail apply --backend com --plan C:\work\mark-read.plan.json --dry-run
white-collar mail apply --backend com --plan C:\work\mark-read.plan.json
```

Before an agent can use these commands, the human owner must run the
human-only grant flow in an interactive terminal, for example:

```powershell
white-collar permissions grant --app mail --backend com --policy read-only --capability mail.metadata.read --target mailbox
white-collar permissions grant --app mail --backend com --policy review --capability mail.body.read --target MESSAGE_ENTRY_ID
white-collar permissions grant --app mail --backend com --policy review --capability mail.write.state --target MESSAGE_ENTRY_ID
white-collar permissions grant --app mail --backend com --policy edit --capability mail.write.compose --target mailbox
white-collar permissions grant --app mail --backend com --policy send --capability mail.write.send --target DRAFT_ENTRY_ID
```

The CLI will refuse a noninteractive permission change and return instructions
for the agent to stop and ask the human.

Mail write plans use `white-collar.plan/v1`, an `id` target, the policy that
matches the operation, and `write.mode: "none"` because the operation changes
the live mailbox rather than creating a file. The first semantic operations are:

```json
{
  "schema": "white-collar.plan/v1",
  "app": "mail",
  "target": {"id": "MESSAGE_ENTRY_ID"},
  "policy": "review",
  "operations": [{"op": "mail_live_mark_read"}],
  "write": {"mode": "none"}
}
```

Use `mail_live_mark_unread` for the other review-level state change. Marking
read or unread requires the human owner to grant `mail.write.state` for the
exact message ID. `mail_live_move` with `{"args":{"folder":"Sent Items"}}`
and `mail_live_delete` are edit-level organization operations and require
`mail.write.organize` for the exact message ID. `--dry-run` resolves and
validates the message but does not modify it.

To create a draft, use an edit-level plan whose target is `mailbox` or an exact
Outlook account address. The operation is bounded to recipients, subject, body,
and optional cc/bcc fields; it saves a draft but never sends it:

```json
{
  "schema": "white-collar.plan/v1",
  "app": "mail",
  "target": {"id": "tgemicioglu@outlook.com"},
  "policy": "edit",
  "operations": [{
    "op": "mail_live_create_draft",
    "args": {
      "to": "person@example.com",
      "subject": "A reviewable draft",
      "body": "Draft body",
      "cc": ""
    }
  }],
  "write": {"mode": "none"}
}
```

Creating a draft requires the human owner to grant `mail.write.compose` for the
mailbox or account target. Sending that draft remains a separate `send`-level
operation and exact draft grant.

To send an existing Outlook draft, use `mail_live_send` with the draft's entry
ID, `policy: "send"`, and a human grant for `mail.write.send`. The send
operation does not compose, reply, or forward, and refuses an item Outlook
already marks as sent.

Search evaluates only message metadata: subject, sender, recipients, and
attachment count. Supported query selectors are `from:`, `to:`, `cc:`,
`subject:`, `has:attachment`, and unqualified terms. It does not search message
bodies. The default folder is `Inbox`; the backend currently supports the
standard default folders `Inbox`, `Sent Items`, `Drafts`, `Deleted Items`,
`Outbox`, and `Junk Email`.

`mail read` returns metadata under `read-only`. Adding `--include-body` is a
sensitive read and requires `review` or `edit`; the permission check happens
before the adapter accesses Outlook. The grant is stored in protected OS
credential storage, not in an agent-editable file. Mail organization is
read/unread state is `review`-level; organization and draft composition are
`edit`-level; sending is `send`-only. All require exact human grants.
Forwarding and attachment-read capabilities remain unavailable in this
milestone.
