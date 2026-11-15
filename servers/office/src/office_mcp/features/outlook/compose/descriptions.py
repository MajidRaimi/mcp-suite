from __future__ import annotations

from office_mcp.shared.owner import Owner
from office_mcp.shared.phrases import (
    FOREGROUND_NOTE,
    KEYS_ONLY_NOTE,
    NO_RETRY_NOTE,
    OUTLOOK_UNAVAILABLE,
)

CHECK_FIRST = "Call office_check_text first when unsure."
RESULT_NOTE = (
    "Returns `status`: `sent` (found in Sent Items), `submitted` (waiting in the Outbox), "
    "`unconfirmed` (Send was pressed but the mail was not found yet) or `dry_run`, plus "
    "`text_sent` (the exact text after persona fixes such as an appended closing) and "
    "`persona` (version, fixes, warnings). `dry_run=true` runs every check and returns the "
    "final text without opening a compose window."
)
SAFETY_NOTE = (
    "Before anything opens, the addresses, the persona rules and the recipient allowlist are "
    "checked; the same mail twice within two minutes, or one already in Sent Items from the "
    "last ten minutes, is refused unless `allow_duplicate=true`. Recipients, subject and text "
    "are read back from the compose window before the single Send press; on any mismatch the "
    "unsent draft is discarded (a copy lands in Deleted Items) and nothing is sent."
)
SENSITIVITY_NOTE = (
    "`sensitivity` picks a label from Outlook's Sensitivity menu before Send, given as the "
    'menu path such as ["General"] or ["Confidential", "Internal"]; the compose must show '
    "that label or nothing is sent. Pass it when Outlook refuses a send because a "
    "sensitivity label is required."
)
ADDRESS_NOTE = (
    "Recipients must be full email addresses (find them with outlook_people); names are "
    "refused. Plain text only; newlines are kept."
)


def outlook_send(owner: Owner, rules: str) -> str:
    return join(
        f"Send a new email from {owner.possessive} Outlook right away: no draft is kept.",
        ADDRESS_NOTE,
        "`to` takes 1 to 20 addresses, `cc` and `bcc` up to 20 each; `subject` 1 to 255 "
        "characters; `body` 1 to 20000 characters.",
        SAFETY_NOTE,
        SENSITIVITY_NOTE,
        RESULT_NOTE,
        NO_RETRY_NOTE,
        FOREGROUND_NOTE,
        OUTLOOK_UNAVAILABLE.strip(),
        rules_tail(rules),
    )


def outlook_reply(owner: Owner, rules: str) -> str:
    return join(
        f"Reply as {owner.label} to the latest message of a conversation and send it right away.",
        KEYS_ONLY_NOTE,
        "Reply goes to the sender; `reply_all=true` also goes to everyone on the message "
        f"except {owner.label}. `cc` and `bcc` add more addresses. The other side's latest "
        "message sets the expected language; `allow_language_mismatch=true` overrides that "
        "check.",
        ADDRESS_NOTE,
        SAFETY_NOTE,
        SENSITIVITY_NOTE,
        RESULT_NOTE,
        "Also returns `in_reply_to` and the `recipients` the reply went to.",
        "Selecting the conversation marks it read.",
        NO_RETRY_NOTE,
        FOREGROUND_NOTE,
        OUTLOOK_UNAVAILABLE.strip(),
        rules_tail(rules),
    )


def outlook_forward(owner: Owner, rules: str) -> str:
    return join(
        f"Forward the latest message of a conversation from {owner.possessive} Outlook and "
        "send it right away.",
        KEYS_ONLY_NOTE,
        "`to` takes 1 to 20 addresses; `body` is an optional note above the forwarded "
        "message and is checked against the persona when not empty.",
        ADDRESS_NOTE,
        SAFETY_NOTE,
        SENSITIVITY_NOTE,
        RESULT_NOTE,
        "Also returns `forwarded` with the key and subject.",
        "Selecting the conversation marks it read.",
        NO_RETRY_NOTE,
        FOREGROUND_NOTE,
        OUTLOOK_UNAVAILABLE.strip(),
        rules_tail(rules),
    )


def outlook_undo_send(owner: Owner) -> str:
    return join(
        "Take back the last mail this server sent, using Outlook's Undo Send, when it was "
        "sent less than 30 seconds ago and Outlook still offers it.",
        "The reopened draft is checked against the subject that was sent and then discarded "
        "(a copy lands in Deleted Items).",
        "Returns `ok`, `undone_subject`, `detail` and `focus_restored`; fails without doing "
        f"anything when there is nothing to undo. Use it only when {owner.label} asks to "
        "stop a mail that was just sent.",
        FOREGROUND_NOTE,
    )


def rules_tail(rules: str) -> str:
    text = rules.strip()
    return text if text.endswith(CHECK_FIRST) else f"{text} {CHECK_FIRST}".strip()


def join(*parts: str) -> str:
    return " ".join(part.strip() for part in parts if part.strip())
