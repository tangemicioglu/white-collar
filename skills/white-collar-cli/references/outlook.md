# Outlook Classic mail

Outlook is the most sensitive adapter. Outlook COM is disabled by default, and
selecting `--backend com` cannot grant access. The human owner must configure
the protected authority first.

## Read metadata

```powershell
white-collar mail search --backend com --folder Inbox --query "from:person@example.com roadmap" --limit 10
white-collar mail read --backend com --id MESSAGE_ENTRY_ID
```

Search is metadata-only. Supported selectors are `from:`, `to:`, `cc:`,
`subject:`, `has:attachment`, and unqualified terms. The default folder is
`Inbox`; standard folders include `Sent Items`, `Drafts`, `Deleted Items`,
`Outbox`, and `Junk Email`. A normal read returns metadata. Add
`--include-body --policy review` only when the user authorized the sensitive
read and an exact `mail.body.read` grant exists.

## Mail writes

Use a plan for state or organization changes and dry-run first:

```powershell
white-collar mail apply --backend com --plan C:\work\mark-read.plan.json --dry-run
white-collar mail apply --backend com --plan C:\work\mark-read.plan.json
```

Mark read/unread is `review` with `mail.write.state`. Move and delete are
`edit` with `mail.write.organize`. Deleting follows Outlook's normal Deleted
Items behavior. Mail live plans use `target.id` and `write.mode: "none"`:

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

## Draft and send

Draft composition is bounded to account/mailbox, recipients, subject, body,
and optional cc/bcc. It never sends:

```powershell
white-collar mail draft `
  --account tgemicioglu@outlook.com `
  --to person@example.com `
  --subject "Review" `
  --body "Draft body" `
  --dry-run
```

Without dry-run, the shortcut requires `edit` composition authority. Sending is
a separate command and permission level and accepts only an existing draft:

```powershell
white-collar mail send --draft-id DRAFT_ENTRY_ID --dry-run
white-collar mail send --draft-id DRAFT_ENTRY_ID
```

Sending requires `send` and `mail.write.send` for that exact draft. It does not
compose, reply, forward, or send an item already marked sent. Never broaden the
target to `mailbox` to make a send succeed.

Forwarding, replies, attachment reads, and arbitrary Outlook COM operations are
outside this CLI surface in the current milestone.
