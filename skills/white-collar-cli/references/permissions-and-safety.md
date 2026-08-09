# Permissions and safety

The policy in a plan or command expresses the requested risk level. It does
not grant authority. The authority layer checks the operation, backend,
capability, and exact target before an adapter is called.

## Profiles

| Profile | Meaning |
| --- | --- |
| `read-only` | Inspect metadata/content; no writes. |
| `review` | Save-as Office artifacts; read sensitive mail bodies; mark mail read/unread. |
| `edit` | In-place Office writes with a snapshot; organize mail; compose drafts. |
| `send` | Send an existing Outlook draft only. |

Word and PowerPoint have built-in review-level defaults for the ordinary local
and save-as authoring path. In-place edits, screen capture, and some live COM
actions may require additional authority. Outlook COM and its writes are
disabled by default.

## Human owner actions

Only a human owner should run and confirm permission changes:

```powershell
white-collar setup
white-collar setup --preset safe
white-collar setup --preset office-authoring
white-collar setup --preset outlook-review
white-collar setup --preset outlook-send
```

The presets are intentionally bounded:

- `safe`: Word/PowerPoint `review`; Outlook disabled.
- `office-authoring`: Word/PowerPoint `edit`; Outlook disabled.
- `outlook-review`: Outlook review; other applications unchanged.
- `outlook-send`: Outlook send; other applications unchanged.

If a user wants application setup, explain what the selected profile permits and
ask the user to run the interactive command in their own terminal. Do not
create a custom configuration file, edit credential storage, press `y`, or
retry through a broader scope.

For diagnostics, the agent may use:

```powershell
white-collar doctor
white-collar permissions show --redacted
white-collar permissions check --capability mail.body.read --policy review --backend com --target MESSAGE_ID
```

`doctor` and `--redacted` avoid exposing exact owner-grant targets. Use
`permissions show` without redaction only when the human explicitly needs the
details.

## Target scope

Office file grants are exact paths. Mail message grants are exact message IDs.
The `mailbox` target is a broad human-configured scope used for metadata or
draft organization/composition when the user accepts that risk. Never promote
an exact target to `mailbox` to get around a denial.

Grants live in protected OS credential storage. There is no supported plaintext
authority file, `--authority` override, or plan field that grants permission.
The normal agent trust boundary is the CLI's policy and authority check; it is
not a claim that arbitrary code running as the same OS user cannot access that
user's credentials.

## Consequential mail actions

Keep these separate and explicit:

- Body reads can reveal secrets and require an exact `mail.body.read` grant.
- Mark read/unread is `review` and uses `mail.write.state`.
- Move/delete uses `edit` and `mail.write.organize`.
- Draft creation uses `edit` and `mail.write.compose`; it never sends.
- Sending uses `send` and `mail.write.send` for the exact existing draft.

If a command reports `authority_denied`, do not switch backends, alter the
target, or rerun with a stronger policy. Report the missing capability and ask
the human whether they want to perform the grant themselves.
