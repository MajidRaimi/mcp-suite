from __future__ import annotations

import threading
from datetime import UTC
from pathlib import Path
from zoneinfo import ZoneInfo

import pytest
from whatsapp_fakes import NOW, FakeWaha, make_client

from whatsapp_mcp.features.drafts.service import DraftStore
from whatsapp_mcp.features.inbox.service import InboxService
from whatsapp_mcp.features.waha.client import WahaClient
from whatsapp_mcp.shared.errors import ChatNotFoundError, DraftError, WahaUnavailableError


@pytest.fixture
def waha() -> FakeWaha:
    return FakeWaha()


@pytest.fixture
def client(waha: FakeWaha) -> WahaClient:
    return make_client(waha)


@pytest.fixture
def inbox(client: WahaClient) -> InboxService:
    return InboxService(client, UTC)


def test_needs_reply_lists_direct_chats_waiting_on_me(inbox: InboxService) -> None:
    waiting = inbox.needs_reply(NOW, limit=20, include_groups=False, max_age_days=7)
    assert [item.name for item in waiting] == ["Ahmed Work", "Ahmed"]
    ahmed = waiting[1]
    assert ahmed.unanswered_messages == 2
    assert ahmed.waiting == "1h 0m"
    assert ahmed.last_message == "are you coming?"


def test_needs_reply_can_include_groups(inbox: InboxService) -> None:
    waiting = inbox.needs_reply(NOW, limit=20, include_groups=True, max_age_days=7)
    assert "Family" in [item.name for item in waiting]


def test_read_chat_orders_oldest_first_and_names_senders(inbox: InboxService) -> None:
    summary, lines = inbox.read_chat("15550000001@c.us", 10)
    assert summary.display_name == "Ahmed"
    assert [line.sender for line in lines] == ["me", "Ahmed", "Ahmed"]
    _, group = inbox.read_chat("family", 10)
    assert group[0].sender == "15550000009"


def test_times_are_formatted_in_the_injected_zone(client: WahaClient) -> None:
    _, utc_lines = InboxService(client, UTC).read_chat("Sara", 10)
    _, tokyo_lines = InboxService(client, ZoneInfo("Asia/Tokyo")).read_chat("Sara", 10)
    _, new_york_lines = InboxService(client, ZoneInfo("America/New_York")).read_chat("Sara", 10)
    assert utc_lines[-1].at == "2027-01-15 07:50"
    assert tokyo_lines[-1].at == "2027-01-15 16:50"
    assert new_york_lines[-1].at == "2027-01-15 02:50"


def test_resolve_rejects_ambiguous_and_unknown_names(inbox: InboxService) -> None:
    with pytest.raises(ChatNotFoundError, match="several chats"):
        inbox.resolve("ahm")
    assert inbox.resolve("Ahmed").id == "15550000001@c.us"
    assert inbox.resolve("+1 555 000 0004").display_name == "Ahmed Work"
    with pytest.raises(ChatNotFoundError):
        inbox.resolve("nobody")


def test_search_finds_messages(inbox: InboxService) -> None:
    hits = inbox.search("invoice", NOW, days=7, chat_limit=10)
    assert [(hit.chat, hit.text) for hit in hits] == [("Sara", "invoice attached")]


def test_drafts_send_once_and_restore_on_failure(
    tmp_path: Path, client: WahaClient, waha: FakeWaha
) -> None:
    store = DraftStore(tmp_path / "drafts.json", ttl_seconds=600)
    draft = store.create("15550000001@c.us", "Ahmed", "  on my way  ", NOW)
    assert draft.text == "on my way"
    waha.fail_send = True
    taken = store.take(draft.draft_id, NOW)
    with pytest.raises(WahaUnavailableError):
        client.send_text(taken.chat_id, taken.text)
    store.restore(taken, NOW)
    waha.fail_send = False
    again = store.take(draft.draft_id.upper(), NOW + 10)
    assert client.send_text(again.chat_id, again.text) == "sent-1"
    assert waha.sent == [{"session": "default", "chatId": "15550000001@c.us", "text": "on my way"}]
    with pytest.raises(DraftError):
        store.take(draft.draft_id, NOW + 20)


def test_drafts_expire_and_validate(tmp_path: Path) -> None:
    store = DraftStore(tmp_path / "drafts.json", ttl_seconds=60)
    draft = store.create("c", "C", "hello", NOW)
    with pytest.raises(DraftError, match="expired"):
        store.take(draft.draft_id, NOW + 120)
    with pytest.raises(DraftError, match="empty"):
        store.create("c", "C", "   ", NOW)


def _race_take(store: DraftStore, draft_id: str) -> list[str]:
    barrier = threading.Barrier(2)
    outcomes: list[str] = []

    def take_once() -> None:
        barrier.wait()
        try:
            store.take(draft_id, NOW)
        except DraftError:
            outcomes.append("refused")
        else:
            outcomes.append("taken")

    workers = [threading.Thread(target=take_once) for _ in range(2)]
    for worker in workers:
        worker.start()
    for worker in workers:
        worker.join()
    return sorted(outcomes)


def test_concurrent_take_hands_a_draft_to_exactly_one_caller(tmp_path: Path) -> None:
    store = DraftStore(tmp_path / "drafts.json", ttl_seconds=600)
    for _ in range(50):
        draft = store.create("c", "C", "hello", NOW)
        assert _race_take(store, draft.draft_id) == ["refused", "taken"]


def test_concurrent_creates_keep_every_draft(tmp_path: Path) -> None:
    store = DraftStore(tmp_path / "drafts.json", ttl_seconds=600)
    barrier = threading.Barrier(8)
    created: list[str] = []

    def create_one() -> None:
        barrier.wait()
        created.append(store.create("c", "C", "hello", NOW).draft_id)

    workers = [threading.Thread(target=create_one) for _ in range(8)]
    for worker in workers:
        worker.start()
    for worker in workers:
        worker.join()
    for draft_id in created:
        assert store.take(draft_id, NOW).draft_id == draft_id


def test_client_surfaces_waha_errors() -> None:
    broken = make_client(FakeWaha(), api_key="wrong")
    with pytest.raises(WahaUnavailableError, match="401"):
        broken.chats(10)
