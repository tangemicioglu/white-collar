# white-collar

`white-collar` is an agent-friendly local command-line control plane for a small,
intentional set of Microsoft Office workflows. It gives Word, PowerPoint, and
Outlook adapters shared targeting, policy checks, versioned plans and results,
dry-runs, save-as/snapshot rules, and validation while keeping each application's
commands app-specific.

This repository is an early v0.1 implementation. It is published as source for
review and local use; no Python package release is currently provided.

## What works today

The supported runtime is Windows with desktop Microsoft Office installed. Word,
PowerPoint, and Outlook all use their finite semantic COM adapters; there is no
raw OOXML or local mailbox backend in the public CLI:

* `word inspect` opens a closed document read-only when needed and returns Word
  text, counts, size, and metadata.
* `word create --output ...` creates a new blank, valid `.docx` through Word
  COM without a source document.
* `word apply --plan ...` and `word replace ...` operate through Word's live
  semantic range, formatting, comment, revision, layout, watermark, and save
  operations.
* `slides inspect` and `slides apply --plan ...` operate through PowerPoint COM,
  including native slide rendering with `Slide.Export`.
* `slides create --output ...` creates a new `.pptx` with one blank slide
  through PowerPoint COM without a source presentation.
* `mail search`, `mail read`, and bounded mail writes operate through Outlook
  Classic COM only; Outlook remains disabled until an explicit owner grant.
* Save-as, snapshots, target hashes, dry-runs, policy checks, and validation are
  shared infrastructure around those live Office adapters.

The COM adapters are finite semantic surfaces, not arbitrary COM dispatchers.
Their app-specific operation catalogs are documented in `docs/`.

## Recorded showcase

The [recorded Office showcase](docs/demo/white-collar-office-demo.mp4) runs the
same broad real-operation coverage as the Word and PowerPoint E2E gates. Word
starts as a blank real document and is filled through short CLI chapters; the
same live Word window remains visible while content, formatting, tables,
comments, revisions, layout, metadata, and exports are exercised. PowerPoint
shows the actual affected slide as slides, shapes, charts, notes, sections,
media, motion, templates, and exports are applied. The fixtures are disposable
and synthetic; the edits are performed by the live COM adapters, not by a mock
or a rendered animation.

The recorder uses a bounded display hint for these disposable review copies so
Word can keep the live COM document open after a save-as chapter. This improves
visual continuity without changing normal save-as behavior or permission
defaults. A few Word title changes reflect real COM save-as boundaries; the
document is not closed and reopened for every operation.

The CLI’s own `word_screen_capture` and `slides_screen_capture` operations
remain sensitive, explicit-grant capabilities; the README video uses a local
window recorder instead, so recording does not broaden the Office permission
defaults.

To reproduce it on Windows with desktop Office and `ffmpeg` on `PATH`:

```powershell
python scripts/record-office-demo.py --output docs/demo/white-collar-office-demo.mp4 --force
```

Recording is an optional demonstration tool. The regular CLI does not require
`ffmpeg`.

## Setup

Windows, desktop Word/PowerPoint/Outlook as applicable, and Python 3.11 or newer
are required for live use.

### Install a published release

The release workflow attaches a wheel and source distribution to each version
tag. To install the current `0.1.0` wheel with the Office dependencies, run:

```powershell
py -m pip install "white-collar[office] @ https://github.com/tangemicioglu/white-collar/releases/download/v0.1.0/white_collar-0.1.0-py3-none-any.whl"
white-collar doctor
```

Use a Python virtual environment when the installation must remain isolated.
The command installs the CLI and the optional `pywin32`, `Pillow`, and
`python-docx` dependencies. The latter is used only to create the disposable
bootstrap file for the real Word integration gate; live Office edits still go
through Word COM. Microsoft Office is still a separate Windows prerequisite.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest -q
```

The editable install creates the `white-collar` command. Install the Office
dependencies before using the CLI against real files:

```powershell
python -m pip install -e ".[office]"
```

### Build and publish a release

Update `[project].version` in `pyproject.toml`, commit the change, and create a
matching version tag. The release workflow requires the `v` prefix:

```powershell
git tag v0.1.0
git push origin main --tags
```

The workflow verifies the tag, runs the Office-free test suite, builds and
validates the wheel and source distribution, installs the wheel on Windows,
and creates the GitHub release with both packages attached. It does not run
live Office integration tests because hosted runners do not provide the user's
desktop Office installation.

## Commands

All normal operations and expected failures emit one compact JSON object to
standard output using `white-collar.result/v1`. The human-only permission
grant/revoke path uses readable terminal output when interactive; pass
`--json` there to force the result envelope. The envelope includes
`dry_run: true` only when a mutation was explicitly simulated; ordinary
responses omit that false-valued field to keep agent context small.

```powershell
white-collar word inspect C:\work\brief.docx
white-collar word create --output C:\work\new-brief.docx --dry-run
white-collar word create --output C:\work\new-brief.docx
white-collar word inspect C:\work\brief.docx --backend com --render-dir C:\work\brief-rendered
white-collar word apply --plan C:\work\replace.plan.json --dry-run
white-collar word apply --plan C:\work\remove-draft-watermark.plan.json --dry-run
white-collar word replace --target C:\work\brief.docx --find Draft --replace Final --output C:\work\brief-reviewed.docx
white-collar word apply --plan C:\work\replace.plan.json
white-collar word inspect C:\work\brief.docx --backend com
white-collar word apply --plan C:\work\live.plan.json --backend com

white-collar slides inspect C:\work\deck.pptx
white-collar slides create --output C:\work\new-deck.pptx --dry-run
white-collar slides create --output C:\work\new-deck.pptx
white-collar slides inspect C:\work\deck.pptx --backend com
white-collar slides inspect C:\work\deck.pptx --backend com --render-dir C:\work\deck-rendered
white-collar slides apply --plan C:\work\slides.plan.json --dry-run
white-collar slides apply --plan C:\work\slides.plan.json --backend com

white-collar mail search --query "from:ada@example.com roadmap" --limit 10
white-collar mail search --backend com --folder Inbox --query "from:ada@example.com roadmap" --limit 10
white-collar mail read --id MESSAGE_ID
white-collar mail read --backend com --id MESSAGE_ID
white-collar mail read --id MESSAGE_ID --include-body --policy review
white-collar mail apply --backend com --plan C:\work\mark-read.plan.json --dry-run
white-collar mail apply --backend com --plan C:\work\mark-read.plan.json
white-collar mail draft --account tgemicioglu@outlook.com --to person@example.com --subject "Review" --body "Draft body"
white-collar mail send --draft-id DRAFT_ENTRY_ID --dry-run

white-collar setup
white-collar setup --app mail --policy edit
white-collar setup --preset safe
white-collar setup --preset outlook-review
white-collar doctor
white-collar completions powershell | Out-String | Invoke-Expression

white-collar permissions show
white-collar permissions show --redacted
white-collar permissions show --policy review
white-collar permissions check --capability word.write.save_as --policy review --backend com --target C:\work\brief.docx
white-collar permissions check --capability mail.body.read --policy review --backend com --target MESSAGE_ID

# Human owner only: the agent must ask a human to run and confirm this.
white-collar permissions grant --app word --backend com --policy review --target C:\work\brief.docx --target C:\work\brief-reviewed.docx
white-collar permissions grant --app mail --backend com --policy review --capability mail.metadata.read --target mailbox
white-collar permissions grant --app mail --backend com --policy review --capability mail.body.read --target MESSAGE_ID
white-collar permissions grant --app mail --backend com --policy review --capability mail.write.state --target MESSAGE_ID
white-collar permissions grant --app mail --backend com --policy edit --capability mail.write.organize --target MESSAGE_ID
white-collar permissions grant --app mail --backend com --policy edit --capability mail.write.compose --target mailbox
white-collar permissions grant --app mail --backend com --policy send --capability mail.write.send --target DRAFT_ID
white-collar permissions revoke --app mail --backend com --policy review --capability mail.body.read --target MESSAGE_ID
```

Mail commands and inspect commands default to `read-only`. Mail search and a
message metadata read use `mail.metadata.read`; `mail read --include-body`
requires an explicit `review` or `edit` policy and the `mail.body.read`
capability, plus a matching human-owner grant for that policy. No Outlook COM
write capability is granted by the default authority. Mark read/unread is
`review`-level; moving/deleting mail and creating drafts are `edit`-level; and
sending is `send`-only. Each requires an exact human-owner grant.

## Plans

Mutation plans use `white-collar.plan/v1`, name one target and policy, and contain
only app-approved operations. Office file plans use a path target; mail message
plans use a message-ID target, while draft-composition plans use a mailbox or
account target. The checked-in
[plan schema](schemas/plan-v1.schema.json) is the wire contract.

```json
{
  "schema": "white-collar.plan/v1",
  "app": "word",
  "target": {
    "path": "C:/work/brief.docx",
    "expected_sha256": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"
  },
  "policy": "review",
  "operations": [
    {
      "op": "replace_text",
      "find": "Quarterly Draft",
      "replace": "Quarterly Review",
      "occurrence": "all"
    }
  ],
  "write": {
    "mode": "save-as",
    "path": "C:/work/brief-reviewed.docx"
  }
}
```

Creation plans use the output path as their target, a standalone create
operation, and `write.mode: "create"`:

```json
{
  "schema": "white-collar.plan/v1",
  "app": "word",
  "target": {"path": "C:/work/new-brief.docx"},
  "policy": "review",
  "operations": [{"op": "word_live_create_document"}],
  "write": {"mode": "create"}
}
```

For an in-place edit, use the `edit` policy and name a snapshot explicitly:

```json
{
  "mode": "in-place",
  "snapshot": "C:/work/brief.before-white-collar.docx"
}
```

The schema rejects unknown top-level and operation fields. There is deliberately
no field for a COM method, object path, macro, or arbitrary Office invocation.

## Safety model

| Profile | Reads | Dry-run plans | Create/save-as writes | In-place writes |
| --- | ---: | ---: | ---: | ---: |
| `read-only` | yes | no | no | no |
| `review` | yes | yes | yes | no |
| `edit` | yes | yes | yes | yes, with snapshot |
| `send` | mail only | yes | no | no |

The policy is embedded in every mutation plan and checked before an adapter is
called. `review` is intended for agent workflows that create or produce a new
artifact for human review. `edit` is the only profile that permits replacing the target, and
the plan must provide a distinct snapshot path. A plan's policy is a request;
it is never an authority grant.

The `send` level is Outlook-only and applies only to sending an existing draft;
it is separate from `edit` and requires the `mail.write.send` capability.

The CLI loads owner grants from Windows Credential Manager, not from a JSON
configuration file. There is no plaintext-file fallback, no `--authority` path,
and no checked-in grant file for an agent or human to edit. Word and PowerPoint
have a built-in review-level grant for their COM workflows: reads, dry-runs,
creation, and save-as writes work by default. In-place replacement, screen capture, and
Outlook COM remain disabled until explicitly granted by the human owner.

The permission layer maps finite app operations to a smaller shared capability
vocabulary. `white-collar permissions show` exposes that versioned vocabulary;
`permissions check` checks a requested profile against the active owner grant
without invoking an Office adapter. Plans cannot add grants to themselves.
The built-in Word/PowerPoint review grant is intentionally broad for a smooth
live Office authoring workflow, but it still cannot replace an existing target.
Additional owner grants are exact-target records: file writes require the source and
output/snapshot paths to be granted, and message-body access requires the
specific message target. Outlook COM requires a separate explicit owner grant.

`permissions grant` and `permissions revoke` are deliberately human-owner-only.
In an interactive terminal they show a human-readable summary and use a normal
`[y/N]` confirmation; `--json` is available when a human explicitly needs the
machine-readable form. Noninteractive invocations remain JSON so an agent
receives a structured response telling it to stop and ask the human; it must
not create, edit, confirm, or retry a permission change. The grant payload is
stored as a protected credential blob and is not accepted from command-line
JSON or a plan.

`white-collar setup` is the guided human-owner path for application-level
permission configuration. It can configure one application with
`--app/--policy`, or prompt for Word, PowerPoint, and Outlook in sequence. Word
and PowerPoint offer `disabled`, `read-only`, `review`, and `edit`; Outlook also
offers `send`. The setup writes app-scoped owner grants in one protected update.
For Outlook, `edit` covers mailbox organization and draft composition, while
`send` explicitly covers sending any existing draft in the current mailbox.
Use the lower-level exact-target `permissions grant` command when that scope is
too broad. Word and PowerPoint's built-in review defaults remain available even
when setup removes additional owner grants.
See [the setup guide](docs/permissions-setup.md) for scope details and the
exact-target alternative.

This is an authority boundary for normal CLI and agent operation, not a promise
that arbitrary code already running as the same Windows user cannot access that
user's credentials or automate a confirmation UI. Stronger isolation requires a
separate OS account, broker process, or enterprise endpoint policy.

Outlook Classic is a COM-only backend: install the `office` extra and use
`--backend com` (the only accepted backend). Its narrow search/read behavior and
supported standard folders are documented in [the Outlook COM adapter
catalog](docs/outlook-com-operations.md). The authority layer keeps it disabled
until a human owner enables it.

The COM adapters cover the finite Word and PowerPoint semantic operation
vocabularies documented in [the Word COM operation catalog](docs/word-com-operations.md)
and [the PowerPoint COM operation catalog](docs/powerpoint-com-operations.md).
Mutation plans against an existing file require the target to already be open
in the corresponding Office application. The `word create` and `slides create`
shortcuts use standalone COM creation plans and do not require an existing
target. PowerPoint inspection can open a closed target
read-only, and `slides inspect --backend com --render-dir <dir>` uses
PowerPoint's native `Slide.Export` to write one clean PNG per slide; it does not
require `pdf2image`. Word inspection can likewise open a closed target read-only,
export it through Word's native PDF renderer, and rasterize pages with the system
`pdftoppm` command. All three public adapter paths are COM-backed.

`--dry-run` performs plan, policy, targeting, hash, and authority validation but
does not create an output, snapshot, or mailbox change. The COM adapter reports
the semantic operations it would execute; it does not mutate the live Office
application. A live Word apply fails closed if the target is not open, the
target hash changed, or an output path exists.

## Scope and non-goals

The public model stays small:

* Word, slides, and mail have separate protocols and command trees.
* Shared code handles plans, results, targeting, permissions, policy, and error
  envelopes.
* Backends implement semantic operations such as `replace_text`, `search`, and
  `read`; they do not leak COM object models into plans.

This is not a universal Office automation wrapper, a macro runner, an arbitrary
COM dispatcher, or a high-cardinality MCP server. Existing Office MCP projects
may later serve as optional internal backends, but their many tools are not the
CLI's public command model. No upstream source was copied into this v0.1; in
particular, no `ppt-mcp` source is included.

## Development layout

* `whitecollar/` — CLI, shared policy/schema code, and app adapter contracts.
* `schemas/` — versioned JSON Schema documents for plans and results.
* `tests/fixtures/` — fixture plans.
* `tests/` — Office-free policy/CLI tests plus opt-in real Office COM gates.

Run the complete review gate with:

```powershell
python -m pytest -q
python -m compileall -q whitecollar
```

On Windows with Microsoft Office installed, run the live COM gates as well:

```powershell
python -m pytest -q --run-real-word
python -m pytest -q --run-real-powerpoint
```

That opt-in test starts an isolated Word instance, creates a real `.docx`, and
invokes every registered Word COM semantic operation through the public plan
boundary. It asserts operation-specific postconditions (document text,
formatting, tables, lists, headings, revisions, comments, fields, images,
sections, properties, and deletion), validates every snapshot as a reopenable
Word file, and captures a real Word-window screenshot after each mutation.
It is intentionally separate because the normal policy/CLI test suite uses fake
COM objects and does not require Office.

To retain the live evidence instead of putting it under pytest's temporary
directory, set an artifact directory before running the live gate:

```powershell
$env:WHITE_COLLAR_REAL_WORD_ARTIFACT_DIR = "$pwd\.real-word-artifacts"
python -m pytest -q tests/test_word_com_real.py --run-real-word
```

The ignored artifact directory contains the screenshots, native Word
`SaveCopyAs` snapshots, and the source copies used to validate each capture.

For PowerPoint evidence, set a separate artifact directory and run:

```powershell
$env:WHITE_COLLAR_REAL_POWERPOINT_ARTIFACT_DIR = "$pwd\.real-powerpoint-artifacts"
python -m pytest -q tests/test_slides_com_real.py --run-real-powerpoint
```

That opt-in test creates a disposable real deck and invokes every registered
PowerPoint operation through `Plan` and `PowerPointComAdapter`: inspection,
text operations, slide lifecycle, formatting, shapes, images, backgrounds,
duplication, ordering, templates/layouts, notes, sections, visibility, slide
numbers, tables, charts, SmartArt, media, links, accessibility text, motion,
PDF export, save, and screen capture. It checks each operation's postcondition,
reopens every snapshot in PowerPoint, and retains native PowerPoint-window
screenshots plus PowerPoint-exported slide renders.
The normal review gate does not require Office because it uses fake COM objects
at the adapter boundary. The test uses PowerPoint's native
slide export for render inspection when the optional presentation helper's PDF
rasterizer is unavailable.

The PowerPoint catalog is intentionally semantic and finite. The CLI does not
expose arbitrary COM methods or import `ppt-mcp` source; that project may inform
future behavior, but it is not a public command model or a copied backend.

Local live Office evidence is intentionally not published because Office files
can carry personal author metadata or source-deck content. See the
[artifact note](artifacts/office/README.md) for how to reproduce it locally.

## License

White-Collar is released under the [MIT License](LICENSE).
