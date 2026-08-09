# PowerPoint COM operation coverage

The `--backend com` PowerPoint adapter is the only PowerPoint path and exposes a finite semantic
vocabulary. Plan operations are application-level actions, not PowerPoint COM
method names; there is no arbitrary dispatch escape hatch.

| Area | Operations |
| --- | --- |
| Reading | `slides_live_list_open`, `slides_live_get_info`, `slides_live_get_text`, `slides_live_get_slide_text`, `slides_live_find_text` |
| Text | `slides_live_insert_text`, `slides_live_replace_text`, `slides_live_set_title`, `slides_live_add_textbox`, `slides_live_format_text` |
| Slides | `slides_live_add_slide`, `slides_live_delete_slide`, `slides_live_duplicate_slide`, `slides_live_reorder_slide`, `slides_live_set_slide_size` |
| Graphics | `slides_live_add_shape`, `slides_live_add_image`, `slides_live_set_background` |
| Notes/save/capture | `slides_live_set_notes`, `slides_live_save`, `slides_screen_capture` |

`slides inspect --backend com --render-dir <directory>` is a separate
read-only inspection option, not a plan operation. It opens a closed target
read-only when needed and uses PowerPoint's native `Slide.Export` to write
`slide-1.png`, `slide-2.png`, and so on. Existing output files are never
overwritten.

Example PowerPoint plan operation:

```json
{
  "op": "slides_live_replace_text",
  "args": {
    "find_text": "Draft",
    "replace_text": "Final",
    "replace_all": true
  }
}
```

Use `white-collar slides apply --backend com --plan plan.json` on Windows with
PowerPoint and the optional `office` dependencies installed. The target
presentation must already be open in PowerPoint. A mutation plan must use
PowerPoint's built-in default authority allows `review` with save-as. Mutations
must use `review` with save-as or `edit` with an explicit in-place snapshot. Read
operations use `read-only` and `write.mode: "none"`.

`tests/test_slides_com_real.py` is the live verification harness. It creates a
real presentation, executes every operation in `SLIDES_COM_OPERATIONS`, checks
operation-specific postconditions, reopens every snapshot in PowerPoint, and
captures both the native PowerPoint window and exported slides. The test does
not use fake COM objects. `ppt-mcp` source was not copied; it is only behavioral
inspiration pending a conclusive license review.
