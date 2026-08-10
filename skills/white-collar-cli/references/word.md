# Word

## Inspect

Word operations use the live Word COM adapter on Windows. Microsoft Word must
be installed; mutation targets must already be open in Word.

```powershell
white-collar word inspect C:\work\brief.docx --backend com
```

The COM inspector can open a closed document read-only and render native Word
pages:

```powershell
white-collar word inspect C:\work\brief.docx --backend com --render-dir C:\work\brief-rendered
```

The render directory receives `page-1.png`, `page-2.png`, and so on. Existing
outputs are not overwritten. The renderer uses Word's native PDF export and the
system `pdftoppm`; `pdf2image` is not required by the CLI path.

## Create a blank document

```powershell
white-collar word create --output C:\work\new-brief.docx --dry-run
white-collar word create --output C:\work\new-brief.docx
```

The command creates a valid blank `.docx` through Word COM, refuses an
existing output, and closes the new document after saving. It does not require
an existing source document or an open target.

## Simple replacement

For one replacement, use the bounded shortcut:

```powershell
white-collar word replace `
  --target C:\work\brief.docx `
  --find "Quarterly Draft" `
  --replace "Quarterly Review" `
  --output C:\work\brief-reviewed.docx `
  --dry-run
```

Then repeat without `--dry-run` after reviewing the result. Save-as uses
`review` by default and refuses to overwrite an existing output. In-place use
is more consequential:

```powershell
white-collar word replace `
  --target C:\work\brief.docx `
  --find Draft `
  --replace Final `
  --in-place `
  --snapshot C:\work\brief.before-white-collar.docx `
  --policy edit `
  --dry-run
```

The shortcut supports `--occurrence first|all` and `--backend com` (the only
backend),
`--snapshot`, and `--expected-sha256` where exposed by the installed CLI.

Replacement uses Word's live semantic range and is not an OOXML text-node
rewrite. Inspect the document and use the broader semantic plan operations when
the replacement needs formatting, revisions, or comments preserved explicitly.

## Plans

Use `word apply --plan PLAN.json` for multiple or live semantic operations. A
file plan names `target.path`, an optional `target.expected_sha256`, a policy,
operations, and a write mode. For save-as:

```json
{
  "schema": "white-collar.plan/v1",
  "app": "word",
  "target": {"path": "C:/work/brief.docx"},
  "policy": "review",
  "operations": [{"op": "replace_text", "find": "Draft", "replace": "Final", "occurrence": "all"}],
  "write": {"mode": "save-as", "path": "C:/work/brief-reviewed.docx"}
}
```

For an in-place plan use `policy: "edit"`, `write.mode: "in-place"`, and a
distinct snapshot path. Run `word apply --plan PLAN.json --dry-run` first.

## Live semantic operations

The COM backend has a finite operation catalog, including text editing,
formatting, tables, lists, headings, comments/revisions, layout, metadata,
reading, undo, and screen capture. Names are semantic plan operations, not COM
method names. Examples include `word_live_replace_text`,
`word_live_get_page_text`, `word_live_add_comment`, and
`word_screen_capture`. Consult the repository's `docs/word-com-operations.md`
when an exact operation name or argument shape is needed.

Use `--backend com` only on Windows with the optional `office` extra. Live
mutation plans against existing documents require the document to already be
open in Word; create plans are the exception. The CLI still validates policy,
targeting, hash, and write safety before the adapter.
