from __future__ import annotations

from office_mcp.shared.owner import Owner
from office_mcp.shared.phrases import (
    FOREGROUND_NOTE,
    INJECTION_RULE,
    KEYS_ONLY_NOTE,
    OUTLOOK_UNAVAILABLE,
)

BATCH_NOTE = (
    "Takes up to 25 keys; each one is looked up again right before acting, and the result "
    "lists every key with ok, verified and a detail, plus keys `skipped` when the time ran out."
)
FOLDER_NOTE = (
    "`folder` is only a fallback for keys this server has not listed yet; it opens that folder."
)
UNREAD_NOTE = (
    "Conversations keep their unread state, and neighbors that Outlook marks read on its own "
    "are marked unread again."
)
REFS_HELP = "Keys from outlook_conversations or outlook_search (12 hex characters), 1 to 25."
FOLDER_HELP = "Folder to look in when a key was not listed by this server; usually leave empty."
DESTINATION_HELP = "Folder name or path from outlook_folders, or focused, or other."
READ_HELP = "True marks read, false marks unread."
STATE_HELP = "Flag due bucket, or clear to remove the flag."


def common(owner: Owner) -> str:
    return " ".join(
        (BATCH_NOTE, KEYS_ONLY_NOTE, FOLDER_NOTE, FOREGROUND_NOTE, act_on_request(owner))
    )


def act_on_request(owner: Owner) -> str:
    return f"{INJECTION_RULE} Use only when {owner.label} asked for it."


def outlook_archive(owner: Owner) -> str:
    return (
        f"Move conversations from {owner.possessive} mailbox to Archive. Uses the row's own "
        "Archive action, which does not open the conversation or change its unread state; "
        f"falls back to Message > Archive. {UNREAD_NOTE} {common(owner)}{OUTLOOK_UNAVAILABLE}"
    )


def outlook_delete(owner: Owner) -> str:
    return (
        f"Move conversations from {owner.possessive} mailbox to Deleted Items. Never deletes "
        "permanently: keys in Deleted Items or Recoverable Items are refused, and any alert "
        "that mentions permanent deletion is cancelled and stops the call. The detail says how "
        f"many messages the conversation held. {UNREAD_NOTE} {common(owner)}"
        f"{OUTLOOK_UNAVAILABLE}"
    )


def outlook_move(owner: Owner) -> str:
    return (
        f"Move conversations in {owner.possessive} mailbox to another folder, or between the "
        "Focused and Other tabs with `destination` focused or other (never the Always variants). "
        "The folder must match one folder from outlook_folders exactly by name or path; folders "
        "that share a name are refused. Uses a recent folder in the Move menu when Outlook lists "
        "one, else the Move sheet with one exact match, and never moves again after a failed "
        f"check. {UNREAD_NOTE} {common(owner)}{OUTLOOK_UNAVAILABLE}"
    )


def outlook_mark(owner: Owner) -> str:
    return (
        f"Mark conversations in {owner.possessive} mailbox read or unread and check the result "
        f"in the message list. Already matching ones are left alone. {common(owner)}"
        f"{OUTLOOK_UNAVAILABLE}"
    )


def outlook_flag(owner: Owner) -> str:
    return (
        f"Flag conversations in {owner.possessive} mailbox for follow up with a due bucket "
        "(today, tomorrow, this_week, next_week, no_date) or clear the flag, through the row's "
        "context menu, and check the flag afterwards. Unread conversations are marked unread "
        f"again. {common(owner)}{OUTLOOK_UNAVAILABLE}"
    )
