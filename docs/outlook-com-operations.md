# Outlook Classic adapter

The opt-in Outlook backend is intentionally read-only. It uses the current
user's Outlook Classic MAPI profile through COM and has no send, reply, move,
delete, mark, or attachment-download operation.

## Commands

```powershell
white-collar mail search --backend com --folder Inbox --query "from:ada@example.com roadmap"
white-collar mail search --backend com --folder "Sent Items" --query 'subject:"quarterly review"'
white-collar mail read --backend com --id MESSAGE_ENTRY_ID
white-collar mail read --backend com --id MESSAGE_ENTRY_ID --include-body --policy review
```

Search evaluates only message metadata: subject, sender, recipients, and
attachment count. Supported query selectors are `from:`, `to:`, `cc:`,
`subject:`, `has:attachment`, and unqualified terms. It does not search message
bodies. The default folder is `Inbox`; the backend currently supports the
standard default folders `Inbox`, `Sent Items`, `Drafts`, `Deleted Items`,
`Outbox`, and `Junk Email`.

`mail read` returns metadata under `read-only`. Adding `--include-body` is a
sensitive read and requires `review` or `edit`; the permission check happens
before the adapter accesses Outlook. No v0.1 profile grants mail write or
attachment-read capabilities.
