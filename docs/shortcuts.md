# Human shortcuts

The plan file interface remains the stable agent API. These bounded shortcuts
compile directly into the same versioned `Plan`, policy, authority, adapter, and
result pipeline; they do not expose a second execution surface.

## Word replacement

Create a new output file for review:

```powershell
white-collar word replace `
  --target C:\work\brief.docx `
  --find "Quarterly Draft" `
  --replace "Quarterly Review" `
  --output C:\work\brief-reviewed.docx
```

In-place replacement remains explicit and requires a snapshot:

```powershell
white-collar word replace `
  --target C:\work\brief.docx `
  --find Draft `
  --replace Final `
  --in-place `
  --snapshot C:\work\brief.before-white-collar.docx
```

The default policy is `review` for save-as and `edit` for in-place replacement.
Use `--dry-run` to preview either operation.

## Outlook draft and send

Create a bounded draft through the configured Outlook account:

```powershell
white-collar mail draft `
  --account tgemicioglu@outlook.com `
  --to person@example.com `
  --subject "Review" `
  --body "Draft body"
```

The command requires the edit-level composition permission and returns the new
draft ID. Send only an existing draft:

```powershell
white-collar mail send --draft-id DRAFT_ENTRY_ID
```

The send shortcut requires the separate send permission. Both shortcuts accept
`--dry-run`, and both produce the normal machine-readable result envelope.
