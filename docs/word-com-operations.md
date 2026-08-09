# Word COM operation coverage

The opt-in `--backend com` Word adapter covers the finite live-operation
vocabulary referenced from `word-mcp-live` (MIT, inspected at its `main` branch).
The names are semantic white-collar plan operations; they are not exposed COM
method names and there is no arbitrary dispatch escape hatch.

| Area | Operations |
| --- | --- |
| Editing | `word_live_insert_text`, `word_live_delete_text`, `word_live_replace_text`, `word_live_insert_paragraphs`, `word_live_format_text`, `word_live_add_table`, `word_live_format_table`, `word_live_apply_list`, `word_live_setup_heading_numbering`, `word_live_modify_table`, `word_live_save`, `word_live_toggle_track_changes`, `word_live_insert_image`, `word_live_insert_cross_reference`, `word_live_insert_equation` |
| Reading | `word_live_list_open`, `word_live_get_text`, `word_live_take_snapshot`, `word_live_get_diff`, `word_live_snapshot_status`, `word_live_get_page_text`, `word_live_get_paragraph_format`, `word_live_get_info`, `word_live_find_text`, `word_live_get_undo_history`, `word_live_list_cross_reference_items`, `word_live_diagnose_layout` |
| Comments/revisions | `word_live_get_comments`, `word_live_add_comment`, `word_live_list_revisions`, `word_live_reply_to_comment`, `word_live_resolve_comment`, `word_live_delete_comment`, `word_live_accept_revisions`, `word_live_reject_revisions` |
| Layout | `word_live_set_page_layout`, `word_live_add_header_footer`, `word_live_add_page_numbers`, `word_live_add_section_break`, `word_live_set_paragraph_spacing`, `word_live_add_bookmark`, `word_live_add_watermark` |
| Undo/capture | `word_live_undo`, `word_screen_capture` |
| Metadata | `word_live_set_core_properties` |

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

Use `white-collar word apply --backend com --plan plan.json` on Windows with
Microsoft Word and the optional `office` dependencies installed. The target must
identify a document already open in Word. A mutation plan must use `review` with
save-as or `edit` with an explicit in-place snapshot. The COM adapter creates one
Word undo record per semantic operation.

The adapter is intentionally not the default runtime. The default local backend
continues to use Office-free OOXML for the narrow `replace_text` vertical slice,
which keeps tests and automation usable without Word.

`tests/test_word_com_real.py` is the live verification harness. With Word
installed, `python -m pytest -q --run-real-word` creates a disposable real
document and executes all 45 registered Word operations through the plan and
adapter boundary; it does not use fake COM objects. The harness checks the
postcondition of each operation, reopens every snapshot in Word, and captures
the Word window after every mutating operation. Set
`WHITE_COLLAR_REAL_WORD_ARTIFACT_DIR` to retain those PNG screenshots and
snapshot `.docx` files for review, for example:

```powershell
$env:WHITE_COLLAR_REAL_WORD_ARTIFACT_DIR = "$pwd\.real-word-artifacts"
python -m pytest -q tests/test_word_com_real.py --run-real-word
```
