# Outlook Classic adapter

The opt-in Outlook backend uses the current user's Outlook Classic MAPI profile
through COM. Its finite write catalog currently supports marking a message
read/unread, moving it to a standard folder, and deleting it through Outlook's
normal Deleted Items behavior. It does not send, reply, forward, or download
attachments, and it never exposes arbitrary COM dispatch.

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
```

The CLI will refuse a noninteractive permission change and return instructions
for the agent to stop and ask the human.

Mail write plans use `white-collar.plan/v1`, an `id` target, `edit` policy, and
`write.mode: "none"` because the operation changes the live mailbox rather than
creating a file. The first semantic operations are:

```json
{
  "schema": "white-collar.plan/v1",
  "app": "mail",
  "target": {"id": "MESSAGE_ENTRY_ID"},
  "policy": "edit",
  "operations": [{"op": "mail_live_mark_read"}],
  "write": {"mode": "none"}
}
```

Use `mail_live_mark_unread`, `mail_live_move` with
`{"args":{"folder":"Sent Items"}}`, or `mail_live_delete` for the other
bounded operations. Each requires the human owner to grant
`mail.write.organize` for the exact message ID. `--dry-run` resolves and
validates the message but does not modify it.

Search evaluates only message metadata: subject, sender, recipients, and
attachment count. Supported query selectors are `from:`, `to:`, `cc:`,
`subject:`, `has:attachment`, and unqualified terms. It does not search message
bodies. The default folder is `Inbox`; the backend currently supports the
standard default folders `Inbox`, `Sent Items`, `Drafts`, `Deleted Items`,
`Outbox`, and `Junk Email`.

`mail read` returns metadata under `read-only`. Adding `--include-body` is a
sensitive read and requires `review` or `edit`; the permission check happens
before the adapter accesses Outlook. The grant is stored in protected OS
credential storage, not in an agent-editable file. Mail write is `edit`-only and
requires an exact human grant; send/forward and attachment-read capabilities
remain unavailable in this milestone.
