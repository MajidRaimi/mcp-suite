from __future__ import annotations

from fakes.drivers import fake_runtime, teams_rig
from fakes.p6_chats import (
    MARK_UNREAD_ITEM,
    NEWEST_THEIRS_ID,
    OMAR,
    UnreadDesk,
)
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import teams_menu

from office_mcp.features.teams.reading.factory import build_teams_reading
from office_mcp.features.teams.reading.service import TeamsReadingService


def reading(gateway: FakeTeamsGateway) -> TeamsReadingService:
    return build_teams_reading(fake_runtime(), teams_rig(gateway).driver)


def test_keep_unread_marks_the_newest_message_from_others_unread() -> None:
    gateway = FakeTeamsGateway([])
    desk = UnreadDesk(gateway)

    result = reading(gateway).read("omar", 30, 0, keep_unread=True)

    assert result.restored_unread is True
    assert result.notes == ()
    assert gateway.pressed == [OMAR, MARK_UNREAD_ITEM]
    assert desk.omar_unread is True
    assert [message.text for message in result.conversation.messages] == [
        "draft sent",
        "got it",
        "one question",
    ]
    record = result.to_record()
    assert set(record) == {"chat", "loaded", "returned", "messages", "restored_unread", "notes"}
    assert record["chat"]["thread_id"] == OMAR


def test_without_keep_unread_the_chat_stays_read() -> None:
    gateway = FakeTeamsGateway([])
    desk = UnreadDesk(gateway)

    result = reading(gateway).read("omar", 2, 0, keep_unread=False)

    assert result.restored_unread is False
    assert gateway.pressed == [OMAR]
    assert desk.omar_unread is False
    assert result.to_record()["returned"] == 2
    assert result.to_record()["notes"] == []


def test_a_chat_that_was_already_read_needs_no_restore() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway, omar_unread=False)

    result = reading(gateway).read("omar", 30, 0, keep_unread=True)

    assert result.restored_unread is False
    assert result.notes == ()
    assert gateway.pressed == [OMAR]


def test_missing_menu_item_returns_a_note_instead_of_failing() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway)
    gateway.menus[f"body:{NEWEST_THEIRS_ID}"] = teams_menu("Reply with quote")
    closes_before = gateway.close_calls

    result = reading(gateway).read("omar", 30, 0, keep_unread=True)

    assert result.restored_unread is False
    assert len(result.notes) == 1
    assert result.notes[0].startswith("Teams marked the chat read and it could not be marked")
    assert "no Mark as unread item" in result.notes[0]
    assert gateway.close_calls > closes_before
    assert MARK_UNREAD_ITEM not in gateway.pressed


def test_unverified_restore_returns_a_note() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway, honors_mark=False)

    result = reading(gateway).read("omar", 30, 0, keep_unread=True)

    assert result.restored_unread is False
    assert gateway.pressed == [OMAR, MARK_UNREAD_ITEM]
    assert "does not show the chat as unread" in result.notes[0]


def test_chat_with_only_own_messages_cannot_be_restored() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway, only_mine=True)

    result = reading(gateway).read("omar", 30, 0, keep_unread=True)

    assert result.restored_unread is False
    assert "none from others" in result.notes[0]
    assert gateway.pressed == [OMAR]
