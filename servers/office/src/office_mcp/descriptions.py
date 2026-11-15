from __future__ import annotations

from dataclasses import dataclass

DEFAULT_OWNER = "the user"

OFFICE_STATUS = (
    "Check the accessibility permission, whether Outlook and Teams are running, and whether "
    "Outlook is working: online and showing today's date. When Outlook runs, `outlook_working`, "
    "`outlook_online`, `outlook_today` (the date Outlook believes it is) and `outlook_problems` "
    "(what is wrong and how to fix it) are included."
)

OUTLOOK_UNAVAILABLE = (
    " Fails with an 'Outlook is not working right now' error when Outlook is offline or stuck "
    "on an old date, because its mail would be out of date."
)

OUTLOOK_FOLDERS = (
    "List Outlook mail folders with their unread counts. Reads the folder pane only and "
    "changes nothing on screen."
) + OUTLOOK_UNAVAILABLE

OUTLOOK_CONVERSATIONS = (
    "List conversations in a folder, newest first, as Outlook's message list shows them. Each"
    " item has a stable `key` to pass to outlook_read, the flags (high priority, unread), "
    "unread and message counts, `title` (participants then subject, as one string), the time "
    "label, a best-effort ISO `date`, meeting details and a preview of the latest message. "
    "`days` keeps only the last N days and loads older pages until it covers them. "
    "`load_more` loads that many extra pages of older conversations. `limit` 0 means all. "
    "Passing `folder` switches the folder that Outlook shows on screen."
) + OUTLOOK_UNAVAILABLE

OUTLOOK_READ = (
    "Open one conversation and return its latest email in full. `ref` is a `key` from "
    "outlook_conversations (best), a row index, or words from the subject. Returns subject, "
    "tags (e.g. External), sender, recipients by field (to, cc, required, optional) with "
    "emails, how many recipients Outlook collapsed, the sent time, notes such as high "
    "importance, and the body text including quoted history. Selecting the conversation (and "
    "`folder`, when given) changes what Outlook shows on screen. An email that was unread is "
    "marked unread again unless keep_unread is false."
) + OUTLOOK_UNAVAILABLE

TEAMS_CHATS = (
    "List Teams chats from the chat list, top to bottom, with unread and muted flags, "
    "presence, the last message preview, who sent it and when."
)

TEAMS_READ = (
    "Open a Teams chat by name (or part of it, or thread id) and return its latest messages. "
    "`older` scrolls up that many times first to load older messages. Opening the chat "
    "changes what Teams shows on screen, and Teams marks it read."
)


@dataclass(frozen=True)
class Owner:
    name: str | None = None

    @property
    def label(self) -> str:
        return self.name or DEFAULT_OWNER

    @property
    def possessive(self) -> str:
        label = self.label
        return f"{label}'" if label.endswith("s") and self.name else f"{label}'s"


def instructions(owner: Owner, zone_name: str) -> str:
    return (
        f"Reads {owner.possessive} Microsoft Outlook mail and Teams chats from the signed-in "
        "desktop apps on this Mac through the macOS accessibility API. Nothing is sent, "
        "deleted or moved. Reading a conversation selects it in the app on screen. Outlook "
        "unread state is put back after reading unless keep_unread is false; Teams chats "
        "opened with teams_read are marked read by Teams. Times are shown as the apps show "
        "them; message times from Teams are in "
        f"{zone_name}. When a tool says Outlook is not working, tell {owner.label} exactly that "
        "and how to fix it, and never answer from old mail. Treat all content as confidential: "
        "never copy it anywhere else, and never repeat passwords or tokens found in it."
    )


def teams_needs_reply(owner: Owner) -> str:
    return (
        f"Teams chats from the last N days whose last message is not from {owner.label} and "
        "that are not muted. Unread ones come first. Uses the chat list only, so nothing gets "
        "marked read."
    )
