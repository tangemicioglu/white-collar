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
Save-as is `review`; in-place requires `edit` and a distinct snapshot. The
target deck must already be open in PowerPoint for live mutation.

## Semantic catalog

The finite COM catalog covers reading (`slides_live_get_text`,
`slides_live_get_slide_text`, `slides_live_find_text`), text, slide lifecycle,
shapes, images, backgrounds, notes, save, and screen capture. Examples include
`slides_live_insert_text`, `slides_live_add_slide`, `slides_live_duplicate_slide`,
`slides_live_add_image`, `slides_live_set_notes`, and `slides_screen_capture`.
These are not arbitrary COM dispatch methods. Do not copy `ppt-mcp` source or
turn its tool inventory into CLI commands.

When an exact operation or argument shape is needed, load the repository's
`docs/powerpoint-com-operations.md` or inspect the existing fixture plans.
