from __future__ import annotations

from office_mcp.shared.owner import Owner
from office_mcp.shared.phrases import FOREGROUND_NOTE, KEYS_ONLY_NOTE, OUTLOOK_UNAVAILABLE

VIEW_RESTORED = " Outlook's folder, tab, search and selection are put back afterwards."
UNREAD_RESTORE = (
    " When the conversation was unread it is marked unread again unless keep_unread is "
    "false; that step needs Outlook in front. " + FOREGROUND_NOTE
)

OUTLOOK_FOLDERS = (
    "List Outlook mail folders with `name`, `path` (parents joined by /), `level`, `unread`, "
    "`total` and `unseen` counts where Outlook shows them. Favorites shortcuts are left out. "
    "Reads the folder pane only and changes nothing on screen."
) + OUTLOOK_UNAVAILABLE

OUTLOOK_CONVERSATIONS = (
    "List conversations in a folder, newest first, as Outlook's message list shows them. Each "
    "item has a `key` to pass to outlook_read and the organize and reply tools (valid for the "
    "same day), flags, unread and message counts, `title` (participants then subject), the "
    "time label, a best-effort ISO `date`, meeting details, a preview, `folder`, `view` "
    "(focused or other) and the booleans `flagged`, `replied`, `forwarded`, "
    "`has_attachments` and `pinned`. `view` picks the Focused or Other tab, `both` reads both "
    "and merges them by date, `current` keeps the tab on screen. `days` keeps the last N days "
    "and loads older pages until they are covered; `load_more` loads that many extra pages. "
    "`limit` 0 means all." + VIEW_RESTORED
) + OUTLOOK_UNAVAILABLE

OUTLOOK_READ = (
    "Open one conversation and return its latest email in full. `ref` is a `key` from "
    "outlook_conversations or outlook_search (best), a row index, or words from the subject. "
    "Returns subject, tags, sender, recipients by field with emails, how many recipients "
    "Outlook collapsed, the sent time, notes, the body with quoted history, `messages` (the "
    "thread split into separate emails, newest first, at most `max_messages`) and "
    "`attachments` (name and size). Selecting the conversation marks it read in Outlook."
    + UNREAD_RESTORE
    + VIEW_RESTORED
) + OUTLOOK_UNAVAILABLE

OUTLOOK_SEARCH = (
    "Search mail with Outlook's own search. `query` is passed as typed, so Outlook search "
    "words such as from:, subject: or hasattachments:yes work as far as Outlook supports them. "
    "`scope` is folder, subfolders (both use `folder`, or the folder on screen), mailbox or "
    "all. Returns `{query, scope, count, results}`; each result is a conversation like "
    "outlook_conversations returns, with the `folder` it lives in and the result `group` "
    "(Top Results or All Results). No match gives count 0. The query is added to Outlook's "
    "search history; the search is closed and the previous scope put back afterwards."
    + VIEW_RESTORED
) + OUTLOOK_UNAVAILABLE

OUTLOOK_PEOPLE = (
    "Look up people by name to find their email address before writing to them. Types the "
    "name into Outlook's search box without searching, reads the people suggestions, then "
    "clears the box. Also returns matching people seen in emails read in this session. "
    "Returns `[{name, email, source}]`; when more than one person fits, ask which one is meant "
    "instead of guessing."
) + OUTLOOK_UNAVAILABLE

OUTLOOK_FOLLOW_UPS = (
    "Sent emails from the last `max_days` days that are at least `min_days` old and got no "
    "answer: no conversation with the same subject arrived in the Inbox (both tabs) after "
    "them. Matching is by subject only, so a reply under a new subject is missed. Returns "
    "`[{key, subject, sent_date, to_preview, days_waiting}]`, longest waiting first."
    + VIEW_RESTORED
) + OUTLOOK_UNAVAILABLE

OUTLOOK_SAVE_ATTACHMENT = (
    "Save one attachment of a conversation into the configured download folder and return "
    "`{path, name, size_bytes, notes}`. `name` is the exact attachment name from "
    "outlook_read. Uses the attachment's Download item, waits for the file in Downloads and "
    "moves it into the download folder under a safe, unique name. "
    + KEYS_ONLY_NOTE
    + UNREAD_RESTORE
    + VIEW_RESTORED
) + OUTLOOK_UNAVAILABLE


def outlook_needs_reply(owner: Owner) -> str:
    return (
        f"Mail conversations from the last N days that may be waiting on {owner.label}: not "
        "replied or forwarded, not meeting messages, not drafts and not from automated senders "
        "such as no-reply or notifications. With `view` both, Other tab mail is included only "
        "while unread. Unread ones come first, then newest. Each item is a conversation like "
        "outlook_conversations returns plus `reason`. Uses the message list only, so nothing "
        "gets marked read." + VIEW_RESTORED
    ) + OUTLOOK_UNAVAILABLE
