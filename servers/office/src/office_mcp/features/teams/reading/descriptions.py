from __future__ import annotations

from office_mcp.shared.owner import Owner

TEAMS_CHATS = (
    "List Teams chats from the chat list, top to bottom, with unread and muted flags, "
    "presence, the last message preview, who sent it and when."
)

TEAMS_READ = (
    "Open a Teams chat by name (or part of it, or thread id) and return its latest messages. "
    "`older` scrolls up that many times first to load older messages. Each message has its "
    "author, time and text; a reply that quotes another message carries `quote` (author, "
    "sent, text) with the reply text kept apart, and `edited` tells whether it was edited. "
    "Opening the chat changes what Teams shows on screen, and Teams marks it read. With "
    "keep_unread true, a chat that was unread is marked unread again through the menu of its "
    "newest message from someone else; `restored_unread` says whether that worked and "
    "`notes` say why not."
)

TEAMS_ACTIVITY = (
    "List the Teams activity feed: mentions, replies, reactions, invitations, meeting updates "
    "and missed calls, newest first, with who, what, a preview, the time label and where it "
    "happened. `filter` unread or mentions switches the matching feed filter on for the read "
    "and back afterwards. Side effect: opening Activity makes Teams open the top item, which "
    "can mark it read, and that cannot be undone. Teams is switched back to the app it showed "
    "before."
)

TEAMS_SEARCH = (
    "Search Teams the way the search box suggests while typing: matching messages (sender, "
    "snippet, chat, time), group chats, people, files and query suggestions. Nothing is "
    "opened and nothing is marked read; the search box is cleared afterwards. This is the "
    "quick suggestion list, not the full results page."
)

TEAMS_CHANNELS = (
    "List the teams and the channels Teams shows under them, with unread and private flags "
    "and ids. Collapsed teams are not expanded: they come back once with an empty channel and "
    "`channels_listed` false. Side effect: opening the Teams app makes Teams open the last "
    "channel, which can mark it read. Teams is switched back to the app it showed before."
)

TEAMS_CHANNEL_READ = (
    "Read the latest posts of one channel, given the team and channel names from "
    "teams_channels, with each post's author, time, subject, text, reply count and, when "
    "`replies` is true, the replies Teams shows. Opening the channel marks it read, and "
    "opening the Teams app first makes Teams open the last channel. Teams is switched back to "
    "the app it showed before."
)


def teams_needs_reply(owner: Owner) -> str:
    return (
        f"Teams chats from the last N days whose last message is not from {owner.label} and "
        "that are not muted. Unread ones come first. Uses the chat list only, so nothing gets "
        "marked read."
    )
