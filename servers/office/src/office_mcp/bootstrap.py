from __future__ import annotations

from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol

from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.outlook.mail.service import sent_folder_name
from office_mcp.features.outlook.models import Email, Folder, Summary
from office_mcp.features.outlook.service import MailView
from office_mcp.features.outlook.thread import top_text
from office_mcp.features.persona.bootstrap.collect import collect_corpus
from office_mcp.features.persona.bootstrap.redact import redact
from office_mcp.features.persona.bootstrap.report import require_allowed, write_report
from office_mcp.features.persona.bootstrap.sources import OwnMessage, SentMail
from office_mcp.features.teams.chats.models import Chat
from office_mcp.features.teams.messages.models import Message
from office_mcp.features.teams.reading.service import ChatReading
from office_mcp.shared.runner import Deadline, Runner

MAX_CHAT_READ = 200
CHAT_READ_FACTOR = 3
BOOTSTRAP_BUDGET = 300.0


class MailReader(Protocol):
    def folders(self) -> list[Folder]: ...

    def conversations(
        self,
        folder: str | None,
        unread_only: bool,
        days: int | None,
        limit: int,
        load_more: int = 0,
        view: MailView = MailView.CURRENT,
        deadline: Deadline | None = None,
    ) -> list[Summary]: ...

    def read(self, ref: str, folder: str | None, keep_unread: bool = True) -> Email: ...


class ChatReader(Protocol):
    def chats(self, unread_only: bool, section: str | None, limit: int) -> list[Chat]: ...

    def read(
        self, chat: str, limit: int, older: int, keep_unread: bool, deadline: Deadline | None = None
    ) -> ChatReading: ...


@dataclass
class SkipLog:
    entries: list[str] = field(default_factory=list)

    def add(self, label: str, error: ToolError, names: Sequence[str] = ()) -> None:
        self.entries.append(f"{label}: {redact(str(error), names)}")


def is_own(email: Email, owner_email: str | None) -> bool:
    if owner_email is None or email.sender is None:
        return True
    return email.sender.email.lower() == owner_email.lower()


def sent_mail_of(email: Email) -> SentMail:
    replied_to = email.messages[1].body if len(email.messages) > 1 else ""
    return SentMail(
        subject=email.subject,
        body=top_text(email.body),
        replied_to=replied_to,
        people=mail_people(email),
    )


def mail_people(email: Email) -> tuple[str, ...]:
    sender = (email.sender.name,) if email.sender is not None else ()
    listed = tuple(person.name for people in email.recipients.values() for person in people)
    threaded = tuple(
        part for item in email.messages for part in (item.sender, item.to, item.cc) if part
    )
    return tuple(dict.fromkeys((*sender, *listed, *threaded)))


class SentItemsSource:
    def __init__(
        self,
        runner: Runner,
        mail: Callable[[], MailReader],
        owner_email: Callable[[], str | None],
        skipped: SkipLog,
    ) -> None:
        self._runner = runner
        self._mail = mail
        self._owner_email = owner_email
        self._skipped = skipped

    def sent_mail(self, limit: int) -> Sequence[SentMail]:
        mail = self._mail()
        folder = self._runner.read(lambda _: sent_folder_name(mail.folders()))
        summaries = self._runner.read(
            lambda deadline: mail.conversations(
                folder, False, None, limit, 0, MailView.CURRENT, deadline
            ),
            BOOTSTRAP_BUDGET,
        )
        owner = self._owner_email()
        found: list[SentMail] = []
        for summary in summaries:
            email = self._read(mail, summary.key, folder)
            if email is not None and is_own(email, owner):
                found.append(sent_mail_of(email))
        return found

    def _read(self, mail: MailReader, key: str, folder: str) -> Email | None:
        try:
            return self._runner.read(lambda _: mail.read(key, folder, keep_unread=True))
        except ToolError as error:
            self._skipped.add(f"mail {key}", error)
            return None


def own_messages_of(
    chat: str, messages: Sequence[Message], per_chat: int, names: Sequence[str] = ()
) -> list[OwnMessage]:
    people = tuple(dict.fromkeys((*names, *(message.author for message in messages))))
    found: list[OwnMessage] = []
    context = ""
    for message in messages:
        if not message.mine:
            context = message.text
            continue
        if message.text.strip():
            found.append(
                OwnMessage(chat=chat, text=message.text, replied_to=context, people=people)
            )
    return found[-per_chat:] if per_chat > 0 else []


class ReadChatsSource:
    def __init__(self, runner: Runner, reading: Callable[[], ChatReader], skipped: SkipLog) -> None:
        self._runner = runner
        self._reading = reading
        self._skipped = skipped

    def own_messages(self, chats: int, per_chat: int) -> Sequence[OwnMessage]:
        reading = self._reading()
        rows = self._runner.read(lambda _: reading.chats(False, None, 0))
        picked = [chat for chat in rows if not chat.unread][:chats]
        return [
            item
            for position, chat in enumerate(picked, start=1)
            for item in self._messages(reading, chat, position, per_chat)
        ]

    def _messages(
        self, reading: ChatReader, chat: Chat, position: int, per_chat: int
    ) -> Iterator[OwnMessage]:
        target = chat.thread_id or chat.name
        names = (chat.name, chat.last_sender)
        limit = min(MAX_CHAT_READ, per_chat * CHAT_READ_FACTOR)
        try:
            read = self._runner.read(
                lambda deadline: reading.read(target, limit, 0, False, deadline)
            )
        except ToolError as error:
            self._skipped.add(skipped_chat_label(chat, position), error, names)
            return
        yield from own_messages_of(target, read.conversation.messages, per_chat, names)


def skipped_chat_label(chat: Chat, position: int) -> str:
    return f"chat {chat.thread_id}" if chat.thread_id else f"chat #{position}"


@dataclass(frozen=True)
class BootstrapRequest:
    emails: int
    chats: int
    per_chat: int
    out: Path
    keep_corpus: bool


def run_bootstrap(
    request: BootstrapRequest,
    runner: Runner,
    mail: Callable[[], MailReader],
    reading: Callable[[], ChatReader],
    owner_email: Callable[[], str | None],
) -> dict[str, Any]:
    out = request.out.expanduser()
    require_allowed(out)
    skipped = SkipLog()
    corpus = collect_corpus(
        SentItemsSource(runner, mail, owner_email, skipped),
        ReadChatsSource(runner, reading, skipped),
        request.emails,
        request.chats,
        request.per_chat,
    )
    report = write_report(corpus, out, request.keep_corpus)
    return {
        "report": str(report),
        "emails": len(corpus.emails),
        "messages": len(corpus.messages),
        "skipped": skipped.entries,
    }
