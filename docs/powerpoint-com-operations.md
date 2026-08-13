# PowerPoint COM operation coverage

The `--backend com` PowerPoint adapter is the only PowerPoint path and exposes a finite semantic
vocabulary. Plan operations are application-level actions, not PowerPoint COM
method names; there is no arbitrary dispatch escape hatch.

| Area | Operations |
| --- | --- |
| Reading/inspection | `slides_live_list_open`, `slides_live_get_info`, `slides_live_get_text`, `slides_live_get_slide_text`, `slides_live_find_text`, `slides_live_get_masters`, `slides_live_get_layouts`, `slides_live_get_placeholders`, `slides_live_get_notes`, `slides_live_get_sections`, `slides_live_get_media` |
| Text | `slides_live_insert_text`, `slides_live_replace_text`, `slides_live_set_title`, `slides_live_add_textbox`, `slides_live_format_text` |
| Creation/slides | `slides_live_create_presentation`, `slides_live_add_slide`, `slides_live_delete_slide`, `slides_live_duplicate_slide`, `slides_live_reorder_slide`, `slides_live_set_slide_size`, `slides_live_set_layout`, `slides_live_apply_template`, `slides_live_save_template`, `slides_live_add_section`, `slides_live_delete_section`, `slides_live_set_slide_visibility`, `slides_live_set_slide_numbers` |
| Objects/geometry | `slides_live_add_shape`, `slides_live_add_image`, `slides_live_set_background`, `slides_live_group`, `slides_live_ungroup`, `slides_live_align`, `slides_live_distribute`, `slides_live_z_order`, `slides_live_crop_image`, `slides_live_rotate_shape` |
| Structured/media objects | `slides_live_add_table`, `slides_live_set_table_cell`, `slides_live_add_chart`, `slides_live_add_smartart`, `slides_live_add_media` |
| Links/accessibility/motion | `slides_live_set_hyperlink`, `slides_live_set_alt_text`, `slides_live_set_transition`, `slides_live_add_animation` |
| Notes/output/capture | `slides_live_set_notes`, `slides_live_export_pdf`, `slides_live_save`, `slides_screen_capture` |

`slides inspect --backend com --render-dir <directory>` is a separate
read-only inspection option, not a plan operation. It opens a closed target
read-only when needed and uses PowerPoint's native `Slide.Export` to write
`slide-1.png`, `slide-2.png`, and so on. Existing output files are never
overwritten.

Create a new PowerPoint presentation with one blank slide:

```powershell
white-collar slides create --output C:\work\new-deck.pptx --dry-run
white-collar slides create --output C:\work\new-deck.pptx
```

The create operation uses PowerPoint's `Presentations.Add`, adds one blank
slide, saves through native `SaveAs`, closes the presentation, and refuses an
existing output. It is also available as the standalone plan operation
`slides_live_create_presentation` with `write.mode: "create"`.

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

Template and object operations use application concepts rather than raw COM
calls. Masters and layouts can be inspected by name or index, a slide can be
assigned a layout, and a deck can be saved as a `.potx` through
`slides_live_save_template`. Grouping, alignment, distribution, z-order,
rotation, and image crop accept bounded shape names or indexes. Tables, charts,
SmartArt, and embedded audio/video are created as native PowerPoint objects.
Hyperlinks and alternative text are written to the selected shape; transitions
and effects are written to PowerPoint's slide show and animation timelines.

```json
{
  "op": "slides_live_add_table",
  "args": {
    "slide_index": 2,
    "name": "Results",
    "rows": 2,
    "columns": 2,
    "data": [["Metric", "Value"], ["Pass rate", "100%"]]
  }
}
```

`slides_live_export_pdf` and `slides_live_save_template` write to the plan's
`write.path`, refuse existing outputs, and validate the resulting file. The
PowerPoint adapter uses native `Slide.Export` for slide PNG rendering and
native Office export/copy paths for PDF and template outputs.

Use `white-collar slides apply --backend com --plan plan.json` on Windows with
PowerPoint and the optional `office` dependencies installed. Plans that mutate
an existing presentation must identify it already open in PowerPoint;
standalone create plans are the exception. PowerPoint's built-in default
authority allows `review` creation and save-as. Existing-file mutations must
use `review` with save-as or `edit` with an explicit in-place snapshot. Read
operations use `read-only` and `write.mode: "none"`.

`tests/test_slides_com_real.py` is the live verification harness. It creates a
real presentation, executes every operation in `SLIDES_COM_OPERATIONS`, checks
operation-specific postconditions, reads notes and both embedded audio and
video back through PowerPoint, reopens every snapshot in PowerPoint, and
captures both the native PowerPoint window and exported slides. The test does
not use fake COM objects. `ppt-mcp` source was not copied; it is only behavioral
inspiration pending a conclusive license review.
