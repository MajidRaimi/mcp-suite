from __future__ import annotations

from office_mcp.shared.owner import Owner
from office_mcp.shared.phrases import FOREGROUND_NOTE, NO_RETRY_NOTE, TEAMS_TARGET_NOTE

CHECK_FIRST = "Call office_check_text first when unsure."
OPENS_CHAT = (
    "Opens the chat on screen, so Teams marks it read, and reopens the previous chat after."
)
ALLOWLIST_NOTE = (
    "When OFFICE_MCP_SEND_ONLY_TO is set, chats it does not list are refused before anything "
    "is opened."
)
GROUP_NOTE = (
    "Meeting chats, and chats with more members than the configured limit, are refused unless "
    "group_ok is true; pass it only when the owner asked to write there."
)
TEXT_NOTE = (
    "Plain text up to 4000 characters, sent exactly as checked; line breaks stay in one "
    "message. @mentions are not supported and are sent as typed. A draft already in the "
    "chat's compose box blocks the send instead of being sent or cleared."
)
CHECKS_NOTE = (
    "The persona rules run on the text (with the newest message from someone else as the "
    "language context) and on the text read back from the compose box; the same text sent to "
    "the same chat recently, or matching one of the owner's last 5 messages there, is refused "
    "unless allow_duplicate is true. dry_run runs every check, opens the chat to read the "
    "context, and types nothing."
)
RESULT_NOTE = (
    "Returns `status` sent with `message_ref` set to the new message id once the message "
    "shows in the chat, or unconfirmed when it could not be seen in time."
)
MESSAGE_ID_NOTE = "`message_id` is a message `id` from teams_read."


def sentences(*parts: str) -> str:
    return " ".join(part.strip() for part in parts if part.strip())


def send_rules(rules: str) -> str:
    text = rules.strip()
    return text if text.endswith(CHECK_FIRST) else sentences(text, CHECK_FIRST)


def teams_send(owner: Owner, rules: str) -> str:
    return sentences(
        f"Send a message as {owner.label} in an existing Teams chat. It goes out at once; "
        "there is no draft step.",
        TEAMS_TARGET_NOTE,
        ALLOWLIST_NOTE,
        GROUP_NOTE,
        TEXT_NOTE,
        CHECKS_NOTE,
        RESULT_NOTE,
        OPENS_CHAT,
        FOREGROUND_NOTE,
        NO_RETRY_NOTE,
        send_rules(rules),
    )


def teams_reply(owner: Owner, rules: str) -> str:
    return sentences(
        f"Reply as {owner.label} to one message in an existing Teams chat, quoting it with "
        "Teams' Reply with quote. It goes out at once; there is no draft step.",
        MESSAGE_ID_NOTE,
        TEAMS_TARGET_NOTE,
        ALLOWLIST_NOTE,
        GROUP_NOTE,
        TEXT_NOTE,
        CHECKS_NOTE,
        RESULT_NOTE,
        "Adds `quoted` with the quoted message id and author.",
        OPENS_CHAT,
        FOREGROUND_NOTE,
        NO_RETRY_NOTE,
        send_rules(rules),
    )


def teams_react(owner: Owner) -> str:
    return sentences(
        f"Add or remove {owner.possessive} like, heart, laugh or surprised reaction on one "
        "Teams message. Everyone in the chat sees it.",
        MESSAGE_ID_NOTE,
        TEAMS_TARGET_NOTE,
        ALLOWLIST_NOTE,
        "Reads the reaction first and changes nothing when it is already in the wanted state "
        "(`changed` false). `state` is on or off after the call, `verified` says whether "
        "Teams showed the change.",
        OPENS_CHAT,
        FOREGROUND_NOTE,
    )


def teams_delete_message(owner: Owner) -> str:
    return sentences(
        f"Delete one of {owner.possessive} own Teams messages, for example a message sent by "
        "mistake. Messages from anyone else are refused.",
        MESSAGE_ID_NOTE,
        TEAMS_TARGET_NOTE,
        ALLOWLIST_NOTE,
        "Uses Delete from the message menu and confirms Teams' prompt if one appears; the "
        "result says whether the message disappeared or shows as deleted.",
        OPENS_CHAT,
        FOREGROUND_NOTE,
    )


def teams_mark_unread(owner: Owner) -> str:
    return sentences(
        f"Mark a Teams chat unread so it stays on {owner.possessive} list, using Mark as "
        "unread on the newest message from someone else.",
        TEAMS_TARGET_NOTE,
        "A chat with only the owner's own messages cannot be marked and fails with nothing "
        "changed. Teams may mark the chat read again if it stays open on screen.",
        FOREGROUND_NOTE,
    )
