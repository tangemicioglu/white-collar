# Word COM operation coverage

The `--backend com` Word adapter is the only Word path and covers the finite live-operation
vocabulary referenced from `word-mcp-live` (MIT, inspected at its `main` branch).
The names are semantic white-collar plan operations; they are not exposed COM
method names and there is no arbitrary dispatch escape hatch.

| Area | Operations |
| --- | --- |
| Creation/editing | `word_live_create_document`, `word_live_insert_text`, `word_live_delete_text`, `word_live_replace_text`, `word_live_insert_paragraphs`, `word_live_format_text`, `word_live_apply_style`, `word_live_add_table`, `word_live_format_table`, `word_live_apply_list`, `word_live_setup_heading_numbering`, `word_live_modify_table`, `word_live_save`, `word_live_toggle_track_changes`, `word_live_insert_image`, `word_live_insert_cross_reference`, `word_live_insert_equation`, `word_live_merge_document` |
| Reading/inspection | `word_live_list_open`, `word_live_get_text`, `word_live_take_snapshot`, `word_live_get_diff`, `word_live_snapshot_status`, `word_live_get_page_text`, `word_live_get_paragraph_format`, `word_live_get_info`, `word_live_find_text`, `word_live_get_undo_history`, `word_live_list_cross_reference_items`, `word_live_diagnose_layout`, `word_live_list_styles`, `word_live_list_hyperlinks`, `word_live_list_notes`, `word_live_list_content_controls`, `word_live_get_protection` |
| Links, notes, fields, forms | `word_live_add_hyperlink`, `word_live_remove_hyperlink`, `word_live_add_note`, `word_live_update_fields`, `word_live_insert_toc`, `word_live_set_content_control` |
| Comments/revisions | `word_live_get_comments`, `word_live_add_comment`, `word_live_list_revisions`, `word_live_reply_to_comment`, `word_live_resolve_comment`, `word_live_delete_comment`, `word_live_accept_revisions`, `word_live_reject_revisions` |
| Layout/header/footer | `word_live_set_page_layout`, `word_live_add_header_footer`, `word_live_remove_header_footer`, `word_live_add_page_numbers`, `word_live_add_section_break`, `word_live_set_paragraph_spacing`, `word_live_add_bookmark`, `word_live_add_watermark`, `word_live_remove_watermark` |
| Output/protection | `word_live_export_pdf`, `word_live_set_protection`, `word_live_compare_documents` |
| Undo/capture/metadata | `word_live_undo`, `word_screen_capture`, `word_live_set_core_properties` |

`word inspect --backend com --render-dir <directory>` is a separate read-only
inspection option, not a plan operation. It opens a closed target read-only when
needed, exports through Word's native fixed-format PDF renderer, and writes
`page-1.png`, `page-2.png`, and so on using the system `pdftoppm` command.
Existing output files are never overwritten.

Create a blank Word document with the bounded shortcut:

```powershell
white-collar word create --output C:\work\new-brief.docx --dry-run
white-collar word create --output C:\work\new-brief.docx
```

The create operation uses Word's `Documents.Add` and native `SaveAs2`, closes
the new document after saving, refuses an existing output, and returns a valid
`.docx`. It is also available as the standalone plan operation
`word_live_create_document` with `write.mode: "create"`.

Example plan operation:

```json
{
  "op": "word_live_replace_text",
  "args": {
    "find_text": "ABC Corp",
    "replace_text": "XYZ Ltd",
    "match_case": true,
    "replace_all": true,
    "track_changes": true
  }
}
```

The additional operations remain semantic and bounded. They accept document
ranges, paragraph indexes, bookmarks, or named shapes only where documented;
they do not accept a COM object path or method name.

```json
{
  "op": "word_live_set_content_control",
  "args": {
    "title": "ClientName",
    "value": "Example Client",
    "tag": "client-name",
    "paragraph_index": 1
  }
}
```

`word_live_export_pdf` and `word_live_compare_documents` own their output
path through `write.mode: "save-as"`; they do not overwrite an existing PDF or
comparison document. `word_live_merge_document` inserts a source document at
the target's end and follows the normal snapshot/save-as rules. Protection
types are `none`, `tracked_changes`, `comments`, `forms`, and `read_only`.
`word_live_remove_header_footer` clears authored text and shapes for the
selected header/footer stories. `word_live_update_fields` updates Word story
fields and existing tables of contents.

Remove an exact-text WordArt watermark from header/footer stories. The default
text is `DRAFT`, the default position is both header and footer, and omitting
`section_index` covers all sections. In-place removal still requires an
explicit snapshot under the `edit` policy:

```json
{
  "op": "word_live_remove_watermark",
  "args": {"text": "DRAFT", "position": "both"}
}
```

The adapter temporarily suspends Word revision tracking only while deleting
the matching WordArt objects, then restores the document's prior tracking
setting. Non-matching header/footer content is left untouched.

Use `white-collar word apply --backend com --plan plan.json` on Windows with
Microsoft Word and the optional `office` dependencies installed. Plans that
mutate an existing document must identify it already open in Word; standalone
create plans are the exception. Word's built-in default authority allows
`review` creation and save-as plans. Existing-file mutations must use `review`
with save-as or `edit` with an explicit in-place snapshot. The COM adapter
creates one Word undo record per semantic operation.

This is the only Word adapter exposed by the CLI. The repository's normal tests
use fake COM objects at this boundary so they can run without Word, while the
`--run-real-word` gate exercises the same adapter against a real Word instance.

`tests/test_word_com_real.py` is the live verification harness. With Word
installed, `python -m pytest -q --run-real-word` creates disposable real
documents and executes the existing-file operation matrix plus the standalone
creation operation through the plan and adapter boundary; it does not use fake
COM objects. The harness checks each postcondition and reopens every snapshot
in Word. It also renders representative live mutations through the Word
window; output-producing operations are validated directly as PDF or DOCX
artifacts. Set
`WHITE_COLLAR_REAL_WORD_ARTIFACT_DIR` to retain those PNG screenshots and
snapshot `.docx` files for review, for example:

```powershell
$env:WHITE_COLLAR_REAL_WORD_ARTIFACT_DIR = "$pwd\.real-word-artifacts"
python -m pytest -q tests/test_word_com_real.py --run-real-word
```
