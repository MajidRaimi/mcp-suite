from __future__ import annotations

import threading
from collections.abc import Callable
from typing import Any

from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.container import Services
from office_mcp.features.teams.replies.service import needs_reply
from office_mcp.shared.errors import OfficeError

Record = dict[str, Any]


class OfficeTools:
    def __init__(self, factory: Callable[[], Services]) -> None:
        self._factory = factory
        self._services: Services | None = None
        self._build_lock = threading.Lock()

    def office_status(self) -> Record:
        return self._guarded(lambda s: s.status())

    def outlook_folders(self) -> list[Record]:
        return self._guarded(lambda s: [folder.to_record() for folder in s.outlook.folders()])

    def outlook_conversations(
        self,
        folder: str = "Inbox",
        unread_only: bool = False,
        days: int | None = None,
        limit: int = 50,
        load_more: int = 0,
    ) -> list[Record]:
        return self._guarded(
            lambda s: [
                summary.to_record()
                for summary in s.outlook.conversations(folder, unread_only, days, limit, load_more)
            ]
        )

    def outlook_read(
        self,
        ref: str,
        folder: str | None = None,
        keep_unread: bool = True,
        max_body: int = 20000,
    ) -> Record:
        return self._guarded(
            lambda s: s.outlook.read(ref, folder, keep_unread, max_body).to_record()
        )

    def teams_chats(
        self,
        unread_only: bool = False,
        section: str | None = None,
        limit: int = 0,
    ) -> list[Record]:
        return self._guarded(
            lambda s: [chat.to_record() for chat in s.chats.filtered(unread_only, section, limit)]
        )

    def teams_read(self, chat: str, limit: int = 30, older: int = 0) -> Record:
        return self._guarded(lambda s: s.messages.read(chat, limit, older).to_record())

    def teams_needs_reply(self, days: int = 7, include_meetings: bool = False) -> list[Record]:
        return self._guarded(
            lambda s: [
                chat.to_record()
                for chat in needs_reply(
                    s.chats.filtered(False, None, 0), s.today(), days, include_meetings
                )
            ]
        )

    def _current(self) -> Services:
        with self._build_lock:
            if self._services is None:
                self._services = self._factory()
            return self._services

    def _guarded[T](self, work: Callable[[Services], T]) -> T:
        try:
            current = self._current()
            with current.lock:
                return work(current)
        except OfficeError as error:
            raise ToolError(str(error)) from error
