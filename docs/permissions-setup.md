# Permission setup

`white-collar setup` is the human-owner interface for configuring application
permission bundles. It is intentionally interactive and writes only to the
protected OS credential store. An agent cannot complete setup from a
noninteractive invocation.

Run the guided setup for all applications:

```powershell
white-collar setup
```

Or configure one application directly:

```powershell
white-collar setup --app word --policy edit
white-collar setup --app slides --policy review
white-collar setup --app mail --policy edit
```

For first-run onboarding, use a named preset:

```powershell
white-collar setup --preset safe
white-collar setup --preset office-authoring
white-collar setup --preset outlook-review
white-collar setup --preset outlook-send
```

`safe` selects Word/PowerPoint `review` and leaves Outlook disabled.
`office-authoring` selects Word/PowerPoint `edit` and leaves Outlook disabled.
The two Outlook presets configure only Outlook and preserve settings for the
other applications.

The command displays the proposed application-level change and asks for a
normal `y/N` confirmation. `--json` is available for an explicitly requested
machine-readable result, but it does not remove the interactive human gate.

Profiles are bounded by application:

| Application | Profiles | Scope |
| --- | --- | --- |
| Word | `disabled`, `read-only`, `review`, `edit` | local and COM Word workflows |
| PowerPoint | `disabled`, `read-only`, `review`, `edit` | local and COM PowerPoint workflows |
| Outlook | `disabled`, `read-only`, `review`, `edit`, `send` | current Outlook mailbox |

Outlook setup uses the explicit `mailbox` scope. That scope covers message IDs
in the current mailbox, so `edit` allows organization and draft composition
without granting each message ID separately. `send` also covers sending any
existing draft in that mailbox; the setup summary calls this out because it is
a consequential permission. Use an exact-target grant when mailbox-wide scope
is not appropriate:

```powershell
white-collar permissions grant --app mail --backend com --policy send `
  --capability mail.write.send --target DRAFT_ENTRY_ID
```

Setup replaces owner grants for the applications selected in that run and
preserves owner grants for unselected applications. `disabled` removes owner
grants for that application. Word and PowerPoint still retain their built-in
review-level defaults for local and COM workflows; those defaults are not
removed by setup.
