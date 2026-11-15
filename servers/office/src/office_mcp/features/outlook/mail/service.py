from __future__ import annotations

from collections.abc import Callable, Iterator
from datetime import date
from typing import Any

from office_mcp.features.outlook.conversation import LatestMessage
from office_mcp.features.outlook.mail.attachments import AttachmentService, SavedAttachment
from office_mcp.features.outlook.mail.followups import FollowUp, follow_ups
from office_mcp.features.outlook.mail.search import SearchArea, SearchResult, SearchService
from office_mcp.features.outlook.mail.triage import NeedsReply, mail_needs_reply
from office_mcp.features.outlook.models import Email, Folder, Person, Summary
from office_mcp.features.outlook.people.parse import Contact
from office_mcp.features.outlook.people.service import PeopleService
from office_mcp.features.outlook.reading import DEFAULT_MAX_BODY, DEFAULT_MAX_MESSAGES
from office_mcp.features.outlook.service import MailView, OutlookService
from office_mcp.shared.errors import ToolDisabledError
from office_mcp.shared.runner import Deadline
from office_mcp.shared.text import normalize

INBOX = "Inbox"
SENT_FOLDERS = ("Sent Items", "Sent")
NO_DOWNLOAD_DIR = "saving attachments is off; set OFFICE_MCP_DOWNLOAD_DIR to turn it on"


def participants(email: Email) -> Iterator[Person]:
    if email.sender is not None:
        yield email.sender
    for people in email.recipients.values():
        yield from people


def sent_folder_name(folders: list[Folder]) -> str:
    present = {normalize(folder.name): folder.name for folder in folders}
    return next(
        (present[normalize(name)] for name in SENT_FOLDERS if normalize(name) in present),
        SENT_FOLDERS[0],
    )


class MailService:
    def __init__(
        self,
        core: OutlookService,
        search: SearchService,
        people: PeopleService,
        attachments: AttachmentService | None,
        owner_email: Callable[[], str | None],
        today: Callable[[], date],
    ) -> None:
        self._core = core
        self._search = search
        self._people = people
        self._attachments = attachments
        self._owner_email = owner_email
        self._today = today

    def health_record(self) -> dict[str, Any]:
        return self._core.health_record()

    def folders(self) -> list[Folder]:
        return self._core.folders()

    def conversations(
        self,
        folder: str | None,
        unread_only: bool,
        days: int | None,
        limit: int,
        load_more: int = 0,
        view: MailView = MailView.CURRENT,
        deadline: Deadline | None = None,
    ) -> list[Summary]:
        return self._core.conversations(folder, unread_only, days, limit, load_more, view, deadline)

    def read(
        self,
        ref: str,
        folder: str | None,
        keep_unread: bool = True,
        max_body: int = DEFAULT_MAX_BODY,
        max_messages: int = DEFAULT_MAX_MESSAGES,
    ) -> Email:
        email = self._core.read(ref, folder, keep_unread, max_body, max_messages)
        self._people.remember(participants(email))
        return email

    def latest_message(self, key: str, folder: str | None) -> LatestMessage:
        return self._core.latest_message(key, folder)

    def search(
        self,
        query: str,
        area: SearchArea,
        folder: str | None,
        limit: int,
        load_more: int = 0,
        deadline: Deadline | None = None,
    ) -> SearchResult:
        return self._search.search(query, area, folder, limit, load_more, deadline)

    def people(self, query: str, limit: int) -> list[Contact]:
        return self._people.lookup(query, limit)

    def needs_reply(
        self,
        days: int,
        folder: str,
        view: MailView,
        limit: int,
        deadline: Deadline | None = None,
    ) -> list[NeedsReply]:
        summaries = self._core.conversations(folder, False, days, 0, 0, view, deadline)
        found = mail_needs_reply(summaries, self._owner_email(), self._today(), days)
        return found[: max(limit, 0)]

    def follow_ups(
        self, min_days: int, max_days: int, limit: int, deadline: Deadline | None = None
    ) -> list[FollowUp]:
        sent_folder = sent_folder_name(self._core.folders())
        sent = self._core.conversations(
            sent_folder, False, max_days, 0, 0, MailView.CURRENT, deadline
        )
        inbox = self._core.conversations(INBOX, False, max_days, 0, 0, MailView.BOTH, deadline)
        found = follow_ups(sent, inbox, self._owner_email(), self._today(), min_days, max_days)
        return found[: max(limit, 0)]

    def save_attachment(
        self, ref: str, name: str, folder: str | None, keep_unread: bool
    ) -> SavedAttachment:
        if self._attachments is None:
            raise ToolDisabledError(NO_DOWNLOAD_DIR)
        return self._attachments.save(ref, name, folder, keep_unread)
