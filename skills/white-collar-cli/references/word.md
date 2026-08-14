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

The COM backend has a finite operation catalog. The current groups are:

- Creation/editing: `word_live_create_document`, `word_live_insert_text`,
  `word_live_delete_text`, `word_live_replace_text`,
  `word_live_insert_paragraphs`, `word_live_format_text`,
  `word_live_apply_style`, `word_live_add_table`,
  `word_live_format_table`, `word_live_apply_list`,
  `word_live_setup_heading_numbering`, `word_live_modify_table`,
  `word_live_save`, `word_live_toggle_track_changes`,
  `word_live_insert_image`, `word_live_insert_cross_reference`,
  `word_live_insert_equation`, `word_live_merge_document`.
- Reading and inspection: `word_live_list_open`, `word_live_get_text`,
  `word_live_take_snapshot`, `word_live_get_diff`,
  `word_live_snapshot_status`, `word_live_get_page_text`,
  `word_live_get_paragraph_format`, `word_live_get_info`,
  `word_live_find_text`, `word_live_get_undo_history`,
  `word_live_list_cross_reference_items`, `word_live_diagnose_layout`,
  `word_live_list_styles`, `word_live_list_hyperlinks`,
  `word_live_list_notes`, `word_live_list_content_controls`,
  `word_live_get_protection`.
- Links, notes, fields, and forms: `word_live_add_hyperlink`,
  `word_live_remove_hyperlink`, `word_live_add_note`,
  `word_live_update_fields`, `word_live_insert_toc`,
  `word_live_set_content_control`.
- Comments and revisions: `word_live_get_comments`,
  `word_live_add_comment`, `word_live_list_revisions`,
  `word_live_reply_to_comment`, `word_live_resolve_comment`,
  `word_live_delete_comment`, `word_live_accept_revisions`,
  `word_live_reject_revisions`.
- Layout and metadata: `word_live_set_page_layout`,
  `word_live_add_header_footer`, `word_live_remove_header_footer`,
  `word_live_add_page_numbers`, `word_live_add_section_break`,
  `word_live_set_paragraph_spacing`, `word_live_add_bookmark`,
  `word_live_add_watermark`, `word_live_remove_watermark`,
  `word_live_set_core_properties`.
- Output and control: `word_live_export_pdf`,
  `word_live_set_protection`, `word_live_compare_documents`,
  `word_live_undo`, `word_screen_capture`.

Names are semantic plan operations, not COM method names. Consult the
repository's `docs/word-com-operations.md` when an exact argument shape is
needed. The adapter uses Word's live ranges, stories, tables, fields, and
revisions; it does not rewrite OOXML text nodes.

Use `--backend com` only on Windows with the optional `office` extra. Live
mutation plans against existing documents require the document to already be
open in Word; create plans are the exception. The CLI still validates policy,
targeting, hash, and write safety before the adapter.
