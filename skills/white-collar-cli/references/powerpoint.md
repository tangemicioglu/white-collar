# PowerPoint

## Inspect and render

The slides backend is live PowerPoint COM and requires Windows, PowerPoint, and
the `office` extra:

```powershell
white-collar slides inspect C:\work\deck.pptx --backend com
white-collar slides inspect C:\work\deck.pptx --backend com --render-dir C:\work\deck-rendered
```

The COM inspector can open a closed deck read-only. Rendering uses PowerPoint's
native `Slide.Export` and writes `slide-1.png`, `slide-2.png`, and so on;
`pdf2image` is not required for normal CLI slide viewing. Existing outputs are
not overwritten.

## Create a blank presentation

```powershell
white-collar slides create --output C:\work\new-deck.pptx --dry-run
white-collar slides create --output C:\work\new-deck.pptx
```

The command creates a valid `.pptx` with one blank slide through PowerPoint
COM, refuses an existing output, and closes the new presentation after saving.
It does not require an existing source presentation or an open target.

## Plans

Use `slides apply --plan PLAN.json` for bounded semantic operations:

```json
{
  "schema": "white-collar.plan/v1",
  "app": "slides",
  "target": {"path": "C:/work/deck.pptx"},
  "policy": "review",
  "operations": [{
    "op": "slides_live_replace_text",
    "args": {"find_text": "Draft", "replace_text": "Final", "replace_all": true}
  }],
  "write": {"mode": "save-as", "path": "C:/work/deck-reviewed.pptx"}
}
```

Run `white-collar slides apply --plan PLAN.json --dry-run` before executing.
Save-as and creation are `review`; in-place requires `edit` and a distinct
snapshot. The target deck must already be open in PowerPoint for live mutation
of an existing presentation.

## Native placeholders and semantic text

Text operations address PowerPoint objects semantically. When a slide has
native placeholders, `shape_name` values resolve as follows:

- `Title` selects a title, center-title, or vertical-title placeholder.
- `Body` or `Content` selects a body/content placeholder.
- `Subtitle` selects a subtitle placeholder.

`slides_live_insert_text`, `slides_live_set_title`, and
`slides_live_format_text` therefore edit the built-in field on the selected
slide rather than adding a small overlay textbox. `slides_live_add_slide`
creates a title-and-content (`ppLayoutText`) slide and initializes its native
title field. Use `slides_live_add_textbox` only when a deliberate freeform
textbox is wanted. A plan can still use an exact shape name or positive
`shape_index` for non-placeholder objects.

## Semantic catalog

The finite COM catalog is grouped by behavior:

- Reading: `slides_live_list_open`, `slides_live_get_info`,
  `slides_live_get_text`, `slides_live_get_slide_text`,
  `slides_live_find_text`, `slides_live_get_masters`,
  `slides_live_get_layouts`, `slides_live_get_placeholders`,
  `slides_live_get_notes`, `slides_live_get_sections`,
  `slides_live_get_media`.
- Text: `slides_live_insert_text`, `slides_live_replace_text`,
  `slides_live_set_title`, `slides_live_add_textbox`,
  `slides_live_format_text`.
- Slide lifecycle: `slides_live_create_presentation`,
  `slides_live_add_slide`, `slides_live_delete_slide`,
  `slides_live_duplicate_slide`, `slides_live_reorder_slide`,
  `slides_live_set_slide_size`, `slides_live_set_layout`,
  `slides_live_apply_template`, `slides_live_save_template`,
  `slides_live_add_section`, `slides_live_delete_section`,
  `slides_live_set_slide_visibility`, `slides_live_set_slide_numbers`.
- Objects and geometry: `slides_live_add_shape`, `slides_live_add_image`,
  `slides_live_set_background`, `slides_live_group`,
  `slides_live_ungroup`, `slides_live_align`, `slides_live_distribute`,
  `slides_live_z_order`, `slides_live_crop_image`,
  `slides_live_rotate_shape`.
- Structured/media objects: `slides_live_add_table`,
  `slides_live_set_table_cell`, `slides_live_add_chart`,
  `slides_live_add_smartart`, `slides_live_add_media`.
- Links, accessibility, and motion: `slides_live_set_hyperlink`,
  `slides_live_set_alt_text`, `slides_live_set_transition`,
  `slides_live_add_animation`.
- Output and capture: `slides_live_set_notes`, `slides_live_export_pdf`,
  `slides_live_save`, `slides_screen_capture`.

These are not arbitrary COM dispatch methods. Do not copy `ppt-mcp` source or
turn its tool inventory into CLI commands. For exact argument shapes, load the
repository's `docs/powerpoint-com-operations.md`.
