from __future__ import annotations


MAIL_COM_OPERATIONS = frozenset(
    {
        "mail_live_mark_read",
        "mail_live_mark_unread",
        "mail_live_move",
        "mail_live_delete",
    }
)
MAIL_COM_MUTATING_OPERATIONS = MAIL_COM_OPERATIONS
MAIL_COM_REQUIRED_ARGS = {
    "mail_live_move": {"folder"},
}
