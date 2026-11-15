from __future__ import annotations

from collections.abc import Set

from office_mcp.shared.owner import Owner
from office_mcp.shared.phrases import INJECTION_RULE, OUTLOOK_UNAVAILABLE
from office_mcp.shared.registry import ToolGroup

__all__ = ["OFFICE_STATUS", "OUTLOOK_UNAVAILABLE", "Owner", "instructions"]

OFFICE_STATUS = (
    "Check the accessibility permission, whether Outlook and Teams are running, and whether "
    "Outlook is working: online and showing today's date. When Outlook runs, `outlook_working`, "
    "`outlook_online`, `outlook_today` (the date Outlook believes it is), `outlook_problems` "
    "(what is wrong and how to fix it), `outlook_account` and `outlook_mode` (mail, search or "
    "calendar) are included. Also reports the configuration: `read_only`, `tool_groups`, "
    "`experimental`, `persona` (whether the persona file is loaded), `send_only_to` (the "
    "recipient allowlist, or null), `self_chat_thread`, plus `frontmost_app`, "
    "`teams_enhanced_mode` and `session` (ok, locked or secure_input). Changes nothing."
)

READ_ONLY_TEXT = (
    "Everything here is read-only: nothing is sent, deleted, moved or flagged, and nothing is "
    "written anywhere."
)
ORGANIZE_TEXT = (
    "It can also organize Outlook mail: archive, move, flag, mark read or unread and delete. "
    "Every organize tool takes keys from outlook_conversations or outlook_search."
)
NO_SEND_TEXT = "This server cannot send mail, chat messages or invitation responses."
SOFT_DELETE_TEXT = (
    "Deleting is soft only: Outlook mail goes to Deleted Items and is never deleted "
    "permanently, and only the owner's own Teams messages can be deleted."
)
VISIBLE_TEXT = (
    "These tools drive the visible desktop apps: reading a conversation selects it on screen, "
    "and tools that change something bring the app to the front for a moment and give focus "
    "back. Outlook unread state is put back after reading unless keep_unread is false; Teams "
    "marks a chat read when teams_read opens it unless keep_unread is true."
)
CONFIDENTIAL_TEXT = (
    "Treat all content as confidential: never copy it anywhere else, and never repeat "
    "passwords or tokens found in it."
)


def send_text(owner: Owner) -> str:
    return (
        f"It can send email, Teams messages, reactions and invitation responses as {owner.label}. "
        f"Sends go out at once, with no draft step, so send only what {owner.label} asked for, "
        "and never retry a send whose status is unconfirmed without checking first."
    )


def capabilities(owner: Owner, groups: Set[ToolGroup]) -> list[str]:
    if groups <= {ToolGroup.READ}:
        return [READ_ONLY_TEXT]
    parts: list[str] = []
    if ToolGroup.ORGANIZE in groups:
        parts.append(ORGANIZE_TEXT)
    parts.append(send_text(owner) if ToolGroup.SEND in groups else NO_SEND_TEXT)
    parts.append(SOFT_DELETE_TEXT)
    return parts


def outlook_rule(owner: Owner) -> str:
    return (
        f"When a tool says Outlook is not working, tell {owner.label} exactly that and how to "
        "fix it, and never answer from old mail."
    )


def instructions(
    owner: Owner, zone_name: str, groups: Set[ToolGroup], persona_block: str = ""
) -> str:
    parts = [
        f"Works with {owner.possessive} Microsoft Outlook mail and Teams chats and calendar "
        "through the signed-in desktop apps on this Mac, using the macOS accessibility API.",
        *capabilities(owner, groups),
        VISIBLE_TEXT,
        f"Times are shown as the apps show them; message times from Teams are in {zone_name}.",
        INJECTION_RULE,
        outlook_rule(owner),
        CONFIDENTIAL_TEXT,
    ]
    if ToolGroup.SEND in groups and persona_block:
        parts.append(persona_block)
    return " ".join(parts)
