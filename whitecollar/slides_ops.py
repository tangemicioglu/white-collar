"""Finite semantic PowerPoint live-operation vocabulary."""

SLIDES_COM_OPERATIONS = {
    "slides_live_create_presentation",
    "slides_live_list_open",
    "slides_live_get_info",
    "slides_live_get_text",
    "slides_live_get_slide_text",
    "slides_live_find_text",
    "slides_live_insert_text",
    "slides_live_replace_text",
    "slides_live_add_slide",
    "slides_live_delete_slide",
    "slides_live_set_title",
    "slides_live_add_textbox",
    "slides_live_format_text",
    "slides_live_add_shape",
    "slides_live_add_image",
    "slides_live_set_background",
    "slides_live_duplicate_slide",
    "slides_live_reorder_slide",
    "slides_live_set_notes",
    "slides_live_set_slide_size",
    "slides_live_save",
    "slides_screen_capture",
}

SLIDES_COM_READ_OPERATIONS = {
    "slides_live_list_open",
    "slides_live_get_info",
    "slides_live_get_text",
    "slides_live_get_slide_text",
    "slides_live_find_text",
}

SLIDES_COM_MUTATING_OPERATIONS = SLIDES_COM_OPERATIONS - SLIDES_COM_READ_OPERATIONS

SLIDES_COM_REQUIRED_ARGS = {
    "slides_live_get_slide_text": {"slide_index"},
    "slides_live_find_text": {"search_text"},
    "slides_live_insert_text": {"text"},
    "slides_live_replace_text": {"find_text", "replace_text"},
    "slides_live_delete_slide": {"slide_index"},
    "slides_live_set_title": {"title"},
    "slides_live_add_textbox": {"text"},
    "slides_live_format_text": set(),
    "slides_live_add_shape": set(),
    "slides_live_add_image": {"image_path"},
    "slides_live_set_background": {"color"},
    "slides_live_duplicate_slide": {"slide_index"},
    "slides_live_reorder_slide": {"slide_index", "to_index"},
    "slides_live_set_notes": {"text"},
    "slides_screen_capture": {"output_path"},
}


def is_read_operation(name: str) -> bool:
    return name in SLIDES_COM_READ_OPERATIONS or name == "slides.inspect"
