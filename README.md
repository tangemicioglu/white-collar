# white-collar

`white-collar` is an agent-friendly local command-line control plane for a small,
intentional set of Microsoft Office workflows. It gives Word, PowerPoint, and
Outlook adapters shared targeting, policy checks, versioned plans and results,
dry-runs, save-as/snapshot rules, and validation while keeping each application's
commands app-specific.

This repository is an early, local v0.1 implementation. It is not published as a
package and has no remote repository configured.

## What works today

The Word vertical slice is usable end to end on `.docx` files without Microsoft
Office installed:

* `word inspect` returns document text, counts, size, and SHA-256.
* `word apply --plan ... --dry-run` validates the plan and reports exact matches
  without writing.
* `word apply --plan ...` replaces text in document, header, and footer XML. It
  either writes a new file or edits in place after creating a named snapshot.
* An optional target SHA-256 prevents a stale plan from changing a newer file.
* Existing save-as and snapshot paths are never overwritten.

Text replacement currently matches within individual OOXML text nodes. Word may
split visually continuous text into multiple nodes when formatting changes, so a
phrase crossing such a boundary is reported as unmatched rather than rewritten.

PowerPoint has an opt-in real COM backend with a finite semantic operation
catalog. It requires Windows, PowerPoint, and the optional `office` dependencies;
the default local runtime remains Office-free and returns a structured
`backend_unavailable` result for slides. Outlook is still a narrow, mockable
adapter with an opt-in Outlook Classic COM backend; it has no write
capabilities, and message-body reads require an explicit sensitive-read policy.

## Setup

Python 3.11 or newer is required.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest -q
```

The editable install creates the `white-collar` command. The default Word adapter
uses only the Python standard library; `pywin32` and Microsoft Office are not
required. Install the optional COM dependencies on Windows for the live backend:

```powershell
python -m pip install -e ".[office]"
```

## Commands

All successful operations and expected failures emit one compact JSON object to
standard output using `white-collar.result/v1`. The envelope includes
`dry_run: true` only when a mutation was explicitly simulated; ordinary
responses omit that false-valued field to keep agent context small.

```powershell
white-collar word inspect C:\work\brief.docx
white-collar word inspect C:\work\brief.docx --backend com --render-dir C:\work\brief-rendered
white-collar word apply --plan C:\work\replace.plan.json --dry-run
white-collar word apply --plan C:\work\replace.plan.json
white-collar word inspect C:\work\brief.docx --backend com
white-collar word apply --plan C:\work\live.plan.json --backend com

white-collar slides inspect C:\work\deck.pptx
white-collar slides inspect C:\work\deck.pptx --backend com
white-collar slides inspect C:\work\deck.pptx --backend com --render-dir C:\work\deck-rendered
white-collar slides apply --plan C:\work\slides.plan.json --dry-run
white-collar slides apply --plan C:\work\slides.plan.json --backend com

white-collar mail search --query "from:ada@example.com roadmap" --limit 10
white-collar mail search --backend com --folder Inbox --query "from:ada@example.com roadmap" --limit 10
white-collar mail read --id MESSAGE_ID
white-collar mail read --backend com --id MESSAGE_ID
white-collar mail read --id MESSAGE_ID --include-body --policy review

white-collar permissions show
white-collar permissions show --policy review
white-collar permissions check --capability word.write.save_as --policy review --backend local --target C:\work\brief.docx
white-collar permissions check --capability mail.body.read --policy review --backend com --target MESSAGE_ID

# Human owner only: the agent must ask a human to run and confirm this.
white-collar permissions grant --app word --backend local --policy review --target C:\work\brief.docx --target C:\work\brief-reviewed.docx
white-collar permissions grant --app mail --backend com --policy review --capability mail.metadata.read --target mailbox
white-collar permissions grant --app mail --backend com --policy review --capability mail.body.read --target MESSAGE_ID
white-collar permissions revoke --app mail --backend com --policy review --capability mail.body.read --target MESSAGE_ID
```

Mail commands and inspect commands default to `read-only`. Mail search and a
message metadata read use `mail.metadata.read`; `mail read --include-body`
requires an explicit `review` or `edit` policy and the `mail.body.read`
capability, plus a matching human-owner grant for that policy. No mail write capability is
granted by any v0.1 profile.

## Plans

Mutation plans use `white-collar.plan/v1`, name one target and policy, and contain
only app-approved operations. The checked-in
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

| Profile | Reads | Dry-run plans | Save-as writes | In-place writes |
| --- | ---: | ---: | ---: | ---: |
| `read-only` | yes | no | no | no |
| `review` | yes | yes | yes | no |
| `edit` | yes | yes | yes | yes, with snapshot |

The policy is embedded in every mutation plan and checked before an adapter is
called. `review` is intended for agent workflows that produce a new artifact for
human review. `edit` is the only profile that permits replacing the target, and
the plan must provide a distinct snapshot path. A plan's policy is a request;
it is never an authority grant.

The CLI loads owner grants from protected OS credential storage, not from a
JSON configuration file. On Windows this is the native Windows Credential
Manager; on other platforms an installed OS-keyring backend may be used. There
is no plaintext-file fallback, no `--authority` path, and no checked-in grant
file for an agent or human to edit. With no owner grant, Word and PowerPoint
read-only local/COM inspection remains available and Outlook COM is disabled.

The permission layer maps finite app operations to a smaller shared capability
vocabulary. `white-collar permissions show` exposes that versioned vocabulary;
`permissions check` checks a requested profile against the active owner grant
without invoking an Office adapter. Plans cannot add grants to themselves.
Owner grants are exact-target records: file writes require the source and
output/snapshot paths to be granted, and message-body access requires the
specific message target. The default mail profile can search metadata only
through the unavailable local stub; Outlook COM requires a separate explicit
owner grant.

`permissions grant` and `permissions revoke` are deliberately human-owner-only.
They require an interactive terminal and an exact confirmation phrase. If an
agent reaches this path, it receives a structured response telling it to stop
and ask the human; it must not create, edit, confirm, or retry a permission
change. The grant payload is stored as a protected credential blob and is not
accepted from command-line JSON or a plan.

This is an authority boundary for normal CLI and agent operation, not a promise
that arbitrary code already running as the same Windows user cannot access that
user's credentials or automate a confirmation UI. Stronger isolation requires a
separate OS account, broker process, or enterprise endpoint policy.

Outlook Classic is an opt-in COM backend: install the `office` extra and use
`--backend com`. Its narrow search/read behavior and supported standard folders
are documented in [the Outlook COM adapter catalog](docs/outlook-com-operations.md).
The default backend remains Office-free and returns `backend_unavailable`.

The opt-in COM backends cover the finite Word and PowerPoint semantic operation
vocabularies documented in [the Word COM operation catalog](docs/word-com-operations.md)
and [the PowerPoint COM operation catalog](docs/powerpoint-com-operations.md).
They require the target file to already be open in the corresponding Office
application for mutation plans. PowerPoint inspection can open a closed target
read-only, and `slides inspect --backend com --render-dir <dir>` uses
PowerPoint's native `Slide.Export` to write one clean PNG per slide; it does not
require `pdf2image`. Word inspection can likewise open a closed target read-only,
export it through Word's native PDF renderer, and rasterize pages with the system
`pdftoppm` command. The default backends remain Office-free.

`--dry-run` performs real parsing, targeting, hash checks, and match discovery but
does not create an output or snapshot. A live Word apply fails closed if an
operation matches no text, the target hash changed, or an output path exists.

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
* `tests/` — Office-free unit, integration, and CLI smoke tests.

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
It is intentionally separate because Office is not required for the normal
test suite.

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
duplication, ordering, notes, size, save, and screen capture. It checks each
operation's postcondition, reopens every snapshot in PowerPoint, and retains
native PowerPoint-window screenshots plus PowerPoint-exported slide renders.
The normal review gate does not require Office. The test uses PowerPoint's native
slide export for render inspection when the optional presentation helper's PDF
rasterizer is unavailable.

The PowerPoint catalog is intentionally semantic and finite. The CLI does not
expose arbitrary COM methods or import `ppt-mcp` source; that project may inform
future behavior, but it is not a public command model or a copied backend.

The committed [real Office review artifacts](artifacts/office/README.md) include
one final `.docx` and one final `.pptx` produced by the live operation harness.
