# Command routing

Use the smallest semantic command that matches the requested task. The plan
interface is the stable escape hatch for supported operations that do not have
a bounded shortcut.

## Common routes

| User intent | Command | Default policy |
| --- | --- | --- |
| Read Word text or metadata | `white-collar word inspect FILE.docx` | `read-only` |
| Render Word pages | `white-collar word inspect FILE.docx --backend com --render-dir DIR` | `read-only` |
| Replace simple Word text | `white-collar word replace --target IN --find FIND --replace REPLACE --output OUT` | `review` |
| Inspect a PowerPoint | `white-collar slides inspect FILE.pptx` | `read-only` |
| Render PowerPoint slides | `white-collar slides inspect FILE.pptx --backend com --render-dir DIR` | `read-only` |
| Apply supported Word/PowerPoint semantics | `word apply --plan PLAN` or `slides apply --plan PLAN` | From plan |
| Search Outlook metadata | `white-collar mail search --query QUERY` | `read-only` |
| Read an Outlook message | `white-collar mail read --id MESSAGE_ID` | `read-only` |
| Read a message body | `mail read --id MESSAGE_ID --include-body --policy review` | `review` |
| Create a draft | `mail draft --account ACCOUNT --to TO --subject SUBJECT --body BODY` | `edit` |
| Send an existing draft | `mail send --draft-id DRAFT_ID` | `send` |

All mutations support `--dry-run`. Use it after inspection and before the real
mutation. A dry-run still validates the plan, target, hash, match discovery,
and authority; it must not create a file, snapshot, draft, or mailbox change.

## Shortcut or plan

Use a shortcut when the task is exactly one bounded user-facing action, such as
one Word replacement or a draft with explicit recipients and body. Use a plan
when the task needs multiple operations, a hash guard, an explicit snapshot, a
live semantic operation, or a fixture that should be reviewed before execution.

Shortcuts compile into the same plan, policy, authority, adapter, validation,
and result pipeline. They are not a second permission system.

## Output handling

Normal commands emit one compact JSON result on stdout using
`white-collar.result/v1`; do not request verbose prose just to understand a
failure. `ok: false` is an expected structured result for denied authority,
missing backends, invalid plans, stale hashes, and unmatched edits. Preserve
the error code and details when reporting or retrying.

`permissions grant`, `permissions revoke`, and `setup` are different: in an
interactive terminal they show a human-readable proposal and ask for a normal
confirmation. A noninteractive invocation returns instructions to ask a human.
The agent must not simulate that confirmation.

## Shell examples

```powershell
white-collar doctor
white-collar word inspect C:\work\brief.docx
white-collar word replace --target C:\work\brief.docx --find Draft --replace Final --output C:\work\brief-reviewed.docx --dry-run
white-collar word replace --target C:\work\brief.docx --find Draft --replace Final --output C:\work\brief-reviewed.docx
white-collar slides inspect C:\work\deck.pptx --backend com --render-dir C:\work\deck-rendered
white-collar mail search --query "from:person@example.com roadmap" --limit 10
white-collar mail read --id MESSAGE_ID
```
