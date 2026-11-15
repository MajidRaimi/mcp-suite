from __future__ import annotations

from dataclasses import dataclass, field, replace
from pathlib import Path

import pytest
from fakes.clock import FakeClock
from fakes.lock import MemoryUiLock

from office_mcp.bootstrap import (
    BootstrapRequest,
    ReadChatsSource,
    SentItemsSource,
    SkipLog,
    own_messages_of,
    run_bootstrap,
)
from office_mcp.features.outlook.models import Email, Folder, Person, Summary, ThreadMessage
from office_mcp.features.outlook.service import MailView
from office_mcp.features.teams.chats.models import Chat
from office_mcp.features.teams.messages.models import Conversation, Message
from office_mcp.features.teams.reading.service import ChatReading
from office_mcp.shared.errors import PersonaError, TargetNotFoundError
from office_mcp.shared.runner import Deadline, Runner

OWNER = "alex@example.com"


def runner() -> Runner:
    return Runner(MemoryUiLock(), FakeClock(), 20, 10)


def summary(key: str) -> Summary:
    return Summary(
        key=key,
        index=0,
        group="Today",
        flags=(),
        unread=False,
        unread_count=0,
        message_count=1,
        title="Jane Doe, Project Atlas budget",
        meeting="",
        time_label="9:00 AM",
        date=None,
        preview="",
    )


def email(sender: str, body: str, quoted: str = "") -> Email:
    messages = [ThreadMessage(sender, "", "", "", "Project Atlas budget", body)]
    if quoted:
        messages.append(ThreadMessage("Jane Doe", "", "", "", "Project Atlas budget", quoted))
    return Email(
        subject="Project Atlas budget",
        tags=(),
        sender=Person("Alex Example", sender),
        recipients={},
        hidden_recipients={},
        sent="",
        notes=(),
        body=body,
        truncated=False,
        messages=tuple(messages),
    )


@dataclass
class FakeMail:
    emails: dict[str, Email]
    reads: list[tuple[str, str | None, bool]] = field(default_factory=list)

    def folders(self) -> list[Folder]:
        return [Folder("Inbox", 0), Folder("Sent Items", 0)]

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
        assert folder == "Sent Items"
        return [summary(key) for key in list(self.emails)[:limit]]

    def read(self, ref: str, folder: str | None, keep_unread: bool = True) -> Email:
        self.reads.append((ref, folder, keep_unread))
        if ref == "missing":
            raise TargetNotFoundError("the conversation changed")
        return self.emails[ref]


def chat(thread_id: str, unread: bool) -> Chat:
    return Chat(
        thread_id=thread_id,
        name="Jane Doe",
        kind="chat",
        section="Recent",
        unread=unread,
        muted=False,
        presence="",
        preview="",
        last_sender="",
        from_me=True,
        time_label="",
        last_active=None,
    )


def message(index: int, mine: bool, text: str) -> Message:
    return Message(str(index), "Alex Example" if mine else "Jane Doe", mine, "", text)


@dataclass
class FakeChats:
    rows: list[Chat]
    opened: list[str] = field(default_factory=list)

    def chats(self, unread_only: bool, section: str | None, limit: int) -> list[Chat]:
        return self.rows

    def read(
        self, chat: str, limit: int, older: int, keep_unread: bool, deadline: Deadline | None = None
    ) -> ChatReading:
        self.opened.append(chat)
        messages = [
            message(1, False, "Can you share the numbers?"),
            message(2, True, "Sure, sending them now"),
            message(3, True, "Done"),
        ]
        conversation = Conversation(chat=self.rows[0], messages=messages, loaded=3)
        return ChatReading(conversation=conversation, restored_unread=False)


def test_sent_mail_keeps_only_the_owners_top_text_and_logs_failures() -> None:
    mail = FakeMail(
        {
            "a": email(OWNER, "Numbers attached.\n\nBest regards,", quoted="Please send them"),
            "b": email("jane@example.com", "Thanks Alex"),
            "missing": email(OWNER, "never read"),
        }
    )
    skipped = SkipLog()

    found = SentItemsSource(runner(), lambda: mail, lambda: OWNER, skipped).sent_mail(10)

    assert [item.body for item in found] == ["Numbers attached.\n\nBest regards,"]
    assert found[0].replied_to == "Please send them"
    assert all(keep for _, _, keep in mail.reads)
    assert len(skipped.entries) == 1
    assert skipped.entries[0].startswith("mail missing")


def test_own_messages_skip_unread_chats_and_carry_the_context() -> None:
    chats = FakeChats(
        [chat("19:unread@unq.gbl.spaces", True), chat("19:read@unq.gbl.spaces", False)]
    )

    found = ReadChatsSource(runner(), lambda: chats, SkipLog()).own_messages(5, 10)

    assert chats.opened == ["19:read@unq.gbl.spaces"]
    assert [item.text for item in found] == ["Sure, sending them now", "Done"]
    assert found[0].replied_to == "Can you share the numbers?"


def test_own_messages_keep_the_newest_per_chat() -> None:
    messages = [message(1, True, "one"), message(2, False, "question"), message(3, True, "two")]

    found = own_messages_of("48:notes", messages, 1)

    assert [item.text for item in found] == ["two"]
    assert found[0].replied_to == "question"
    assert own_messages_of("48:notes", messages, 0) == []


def test_bootstrap_checks_the_out_dir_before_reading_anything(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("HOME", str(tmp_path))
    mail = FakeMail({"a": email(OWNER, "hello")})
    request = BootstrapRequest(1, 0, 0, tmp_path / "elsewhere", keep_corpus=False)

    with pytest.raises(PersonaError, match="only inside"):
        run_bootstrap(request, runner(), lambda: mail, lambda: FakeChats([]), lambda: OWNER)

    assert mail.reads == []


def test_bootstrap_writes_the_stats_report(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HOME", str(tmp_path))
    mail = FakeMail({"a": email(OWNER, "Numbers attached.")})
    chats = FakeChats([chat("19:read@unq.gbl.spaces", False)])
    request = BootstrapRequest(5, 2, 5, Path("~/.config/office-mcp"), keep_corpus=False)

    result = run_bootstrap(request, runner(), lambda: mail, lambda: chats, lambda: OWNER)

    report = Path(str(result["report"]))
    assert report.parent == tmp_path / ".config" / "office-mcp"
    assert report.stat().st_mode & 0o777 == 0o600
    assert (result["emails"], result["messages"], result["skipped"]) == (1, 2, [])


def test_skipped_chats_never_log_the_chat_name() -> None:
    nameless = replace(chat("", False), name="Alice Example")
    failing = FailingChats([nameless])
    skipped = SkipLog()

    ReadChatsSource(runner(), lambda: failing, skipped).own_messages(5, 10)

    assert skipped.entries == ["chat #1: no chat named '[name] [name]'"]


def test_own_messages_carry_the_chat_people_for_redaction() -> None:
    messages = [message(1, False, "question"), message(2, True, "answer")]

    found = own_messages_of("19:a@unq.gbl.spaces", messages, 5, ("Alice Example",))

    assert "Alice Example" in found[0].people


@dataclass
class FailingChats(FakeChats):
    def read(
        self, chat: str, limit: int, older: int, keep_unread: bool, deadline: Deadline | None = None
    ) -> ChatReading:
        raise TargetNotFoundError(f"no chat named {chat!r}")
