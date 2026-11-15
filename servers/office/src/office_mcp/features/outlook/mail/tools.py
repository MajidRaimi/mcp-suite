from __future__ import annotations

from collections.abc import Callable
from typing import Annotated, Any, Literal

from pydantic import Field

from office_mcp.features.outlook.mail import descriptions as text
from office_mcp.features.outlook.mail.search import SearchArea
from office_mcp.features.outlook.mail.service import MailService
from office_mcp.features.outlook.service import MailView
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    ToolSpec,
)
from office_mcp.shared.runner import Deadline, Lazy, Runner
from office_mcp.shared.settings import Settings

Record = dict[str, Any]
ViewName = Literal["current", "focused", "other", "both"]
ScopeName = Literal["folder", "subfolders", "mailbox", "all"]
Ref = Annotated[str, Field(min_length=1, max_length=512)]
FolderName = Annotated[str, Field(min_length=1, max_length=256)]


class MailTools:
    def __init__(
        self, runner: Runner, mail: Lazy[MailService], owner: Owner, settings: Settings
    ) -> None:
        self._runner = runner
        self._mail = mail
        self._owner = owner
        self._settings = settings

    def specs(self) -> list[ToolSpec]:
        specs = [
            self._read_spec(self.outlook_folders, text.OUTLOOK_FOLDERS),
            self._read_spec(self.outlook_conversations, text.OUTLOOK_CONVERSATIONS),
            self._read_spec(self.outlook_read, text.OUTLOOK_READ, Effect.MARKS_READ),
            self._read_spec(self.outlook_search, text.OUTLOOK_SEARCH),
            self._read_spec(self.outlook_people, text.OUTLOOK_PEOPLE),
            self._read_spec(self.outlook_needs_reply, text.outlook_needs_reply(self._owner)),
            self._read_spec(self.outlook_follow_ups, text.OUTLOOK_FOLLOW_UPS),
        ]
        if self._settings.download_dir is not None:
            specs.append(
                ToolSpec(
                    name="outlook_save_attachment",
                    function=self.outlook_save_attachment,
                    description=EXPERIMENTAL_PREFIX + text.OUTLOOK_SAVE_ATTACHMENT,
                    group=ToolGroup.ORGANIZE,
                    effect=Effect.WRITES_LOCAL,
                    stability=Stability.EXPERIMENTAL,
                )
            )
        return specs

    def outlook_folders(self) -> list[Record]:
        return self._read(lambda mail, _: [folder.to_record() for folder in mail.folders()])

    def outlook_conversations(
        self,
        folder: FolderName = "Inbox",
        unread_only: bool = False,
        days: Annotated[int | None, Field(ge=1, le=365)] = None,
        limit: Annotated[int, Field(ge=0, le=500)] = 50,
        load_more: Annotated[int, Field(ge=0, le=20)] = 0,
        view: ViewName = "current",
    ) -> list[Record]:
        return self._read(
            lambda mail, deadline: [
                summary.to_record()
                for summary in mail.conversations(
                    folder, unread_only, days, limit, load_more, MailView(view), deadline
                )
            ]
        )

    def outlook_read(
        self,
        ref: Ref,
        folder: FolderName | None = None,
        keep_unread: bool = True,
        max_body: Annotated[int, Field(ge=1, le=100000)] = 20000,
        max_messages: Annotated[int, Field(ge=1, le=100)] = 20,
    ) -> Record:
        return self._read(
            lambda mail, _: mail.read(ref, folder, keep_unread, max_body, max_messages).to_record()
        )

    def outlook_search(
        self,
        query: Annotated[str, Field(min_length=1, max_length=256)],
        scope: ScopeName = "mailbox",
        folder: FolderName | None = None,
        limit: Annotated[int, Field(ge=1, le=200)] = 25,
        load_more: Annotated[int, Field(ge=0, le=10)] = 0,
    ) -> Record:
        return self._read(
            lambda mail, deadline: mail.search(
                query, SearchArea(scope), folder, limit, load_more, deadline
            ).to_record()
        )

    def outlook_people(
        self,
        query: Annotated[str, Field(min_length=2, max_length=100)],
        limit: Annotated[int, Field(ge=1, le=20)] = 8,
    ) -> list[Record]:
        return self._read(
            lambda mail, _: [contact.to_record() for contact in mail.people(query, limit)]
        )

    def outlook_needs_reply(
        self,
        days: Annotated[int, Field(ge=1, le=60)] = 7,
        folder: FolderName = "Inbox",
        view: ViewName = "both",
        limit: Annotated[int, Field(ge=1, le=100)] = 30,
    ) -> list[Record]:
        return self._read(
            lambda mail, deadline: [
                item.to_record()
                for item in mail.needs_reply(days, folder, MailView(view), limit, deadline)
            ]
        )

    def outlook_follow_ups(
        self,
        min_days: Annotated[int, Field(ge=0, le=30)] = 2,
        max_days: Annotated[int, Field(ge=1, le=60)] = 14,
        limit: Annotated[int, Field(ge=1, le=100)] = 30,
    ) -> list[Record]:
        return self._read(
            lambda mail, deadline: [
                item.to_record() for item in mail.follow_ups(min_days, max_days, limit, deadline)
            ]
        )

    def outlook_save_attachment(
        self,
        ref: Ref,
        name: Annotated[str, Field(min_length=1, max_length=512)],
        folder: FolderName | None = None,
        keep_unread: bool = True,
    ) -> Record:
        return self._runner.write(
            lambda _: self._mail.get().save_attachment(ref, name, folder, keep_unread).to_record()
        )

    def _read_spec(
        self,
        function: Callable[..., object],
        description: str,
        effect: Effect = Effect.READS,
    ) -> ToolSpec:
        return ToolSpec(
            name=getattr(function, "__name__", ""),
            function=function,
            description=description,
            group=ToolGroup.READ,
            effect=effect,
        )

    def _read[T](self, work: Callable[[MailService, Deadline], T]) -> T:
        return self._runner.read(lambda deadline: work(self._mail.get(), deadline))
