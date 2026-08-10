"""The finite semantic Word live-operation vocabulary.

Names are intentionally app-level operations. They are not COM method names and
cannot be extended by putting an arbitrary method string in a plan.
"""

WORD_COM_OPERATIONS = {
    "word_live_create_document",
    "word_live_insert_text",
    "word_live_delete_text",
    "word_live_replace_text",
    "word_live_insert_paragraphs",
    "word_live_format_text",
    "word_live_add_table",
    "word_live_format_table",
    "word_live_apply_list",
    "word_live_setup_heading_numbering",
    "word_live_modify_table",
    "word_live_save",
    "word_live_toggle_track_changes",
    "word_live_insert_image",
    "word_live_insert_cross_reference",
    "word_live_insert_equation",
    "word_live_list_open",
    "word_live_get_text",
    "word_live_take_snapshot",
    "word_live_get_diff",
    "word_live_snapshot_status",
    "word_live_get_page_text",
    "word_live_get_paragraph_format",
    "word_live_get_info",
    "word_live_find_text",
    "word_live_get_undo_history",
    "word_live_list_cross_reference_items",
    "word_live_diagnose_layout",
    "word_live_get_comments",
    "word_live_add_comment",
    "word_live_list_revisions",
    "word_live_reply_to_comment",
    "word_live_resolve_comment",
    "word_live_delete_comment",
    "word_live_accept_revisions",
    "word_live_reject_revisions",
    "word_live_set_page_layout",
    "word_live_add_header_footer",
    "word_live_add_page_numbers",
    "word_live_add_section_break",
    "word_live_set_paragraph_spacing",
    "word_live_add_bookmark",
    "word_live_add_watermark",
    "word_live_remove_watermark",
    "word_live_undo",
    "word_screen_capture",
    # Present in the reference implementation and useful for metadata control.
    "word_live_set_core_properties",
}

WORD_COM_READ_OPERATIONS = {
    "word_live_list_open",
    "word_live_get_text",
    "word_live_take_snapshot",
    "word_live_get_diff",
    "word_live_snapshot_status",
    "word_live_get_page_text",
    "word_live_get_paragraph_format",
    "word_live_get_info",
    "word_live_find_text",
    "word_live_get_undo_history",
    "word_live_list_cross_reference_items",
    "word_live_diagnose_layout",
    "word_live_get_comments",
    "word_live_list_revisions",
}

WORD_COM_MUTATING_OPERATIONS = WORD_COM_OPERATIONS - WORD_COM_READ_OPERATIONS

WORD_COM_REQUIRED_ARGS = {
    "word_live_insert_text": {"text"},
    "word_live_replace_text": {"find_text", "replace_text"},
    "word_live_insert_paragraphs": {"paragraphs"},
    "word_live_add_table": {"rows"},
    "word_live_modify_table": set(),
    "word_live_insert_image": {"image_path"},
    "word_live_insert_cross_reference": set(),
    "word_live_insert_equation": {"equation"},
    "word_live_find_text": set(),
    "word_live_add_comment": set(),
    "word_live_reply_to_comment": set(),
    "word_live_add_bookmark": {"bookmark_name"},
    "word_live_remove_watermark": set(),
    "word_screen_capture": {"output_path"},
}


def is_read_operation(name: str) -> bool:
    return name in WORD_COM_READ_OPERATIONS or name == "word.inspect"
