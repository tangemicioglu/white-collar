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

PowerPoint and Outlook have narrow, mockable adapter contracts. Their CLI commands
and policy behavior are tested with fake adapters; the default local runtime
returns a structured `backend_unavailable` result until an app-specific backend
is configured.

## Setup

Python 3.11 or newer is required.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest -q
```

The editable install creates the `white-collar` command. The Word adapter uses
only the Python standard library; `pywin32` and Microsoft Office are not required.

## Commands

All successful operations and expected failures emit one compact JSON object to
standard output using `white-collar.result/v1`.

```powershell
white-collar word inspect C:\work\brief.docx
white-collar word apply --plan C:\work\replace.plan.json --dry-run
white-collar word apply --plan C:\work\replace.plan.json

white-collar slides inspect C:\work\deck.pptx
white-collar slides apply --plan C:\work\slides.plan.json --dry-run

white-collar mail search --query "from:ada@example.com roadmap" --limit 10
white-collar mail read --id MESSAGE_ID
```

Mail commands default to `read-only`. Inspect commands also default to
`read-only`; `--policy review` and `--policy edit` are available when a caller
needs to make its granted capability explicit.

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
the plan must provide a distinct snapshot path. Mail is read-only in v0.1.

`--dry-run` performs real parsing, targeting, hash checks, and match discovery but
does not create an output or snapshot. A live Word apply fails closed if an
operation matches no text, the target hash changed, or an output path exists.

## Scope and non-goals

The public model stays small:

* Word, slides, and mail have separate protocols and command trees.
* Shared code handles plans, results, targeting, policy, and error envelopes.
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
