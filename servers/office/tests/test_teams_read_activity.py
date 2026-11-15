from __future__ import annotations

from fakes.drivers import fake_runtime, teams_rig
from fakes.p6_trees import activity_feed, activity_item, rail_handle, teams_screen
from fakes.teams_gateway import FakeTeamsGateway

from office_mcp.features.teams.activity.models import ActivityFilter, ActivityKind
from office_mcp.features.teams.activity.parse import classify, parse_activity
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.features.teams.reading.factory import build_teams_reading
from office_mcp.features.teams.reading.service import TeamsReadingService
from office_mcp.shared.node import Node

OMAR_THREAD = "19:omar@unq.gbl.spaces"
MENTION = activity_item(
    "1", "Omar Haddad mentioned you", "can you check the budget", "2:05 PM", "Project Atlas"
)
REPLY = activity_item("2", "Jane Doe replied", "done", "11/12", "General")
CALL = activity_item("3", "Missed call from Alex Example", "", "Yesterday", "")


class ActivityDesk:
    def __init__(self, gateway: FakeTeamsGateway, unread: bool = False) -> None:
        self.gateway = gateway
        self.app = TeamsApp.CHAT
        self.unread = unread
        self.mentions = False
        gateway.on_press = self.press
        gateway.show(self.render())

    def render(self) -> Node:
        if self.app is TeamsApp.CHAT:
            return teams_screen(TeamsApp.CHAT, title="Chat | Omar Haddad", thread=OMAR_THREAD)
        items = [MENTION, REPLY, CALL]
        if self.unread:
            items = [MENTION]
        if self.mentions:
            items = [item for item in items if item is MENTION]
        feed = activity_feed(items, unread=self.unread, mentions=self.mentions)
        return teams_screen(TeamsApp.ACTIVITY, feed, title="Activity | Omar Haddad")

    def press(self, node: Node) -> None:
        if node.handle == rail_handle(TeamsApp.ACTIVITY):
            self.app = TeamsApp.ACTIVITY
        elif node.handle == rail_handle(TeamsApp.CHAT):
            self.app = TeamsApp.CHAT
        elif node.handle == "toggle:Unread":
            self.unread = not self.unread
        elif node.handle == "toggle:@Mentions":
            self.mentions = not self.mentions
        self.gateway.show(self.render())


def reading(gateway: FakeTeamsGateway) -> TeamsReadingService:
    return build_teams_reading(fake_runtime(), teams_rig(gateway).driver)


def test_titles_map_to_kinds_and_people() -> None:
    cases = {
        "Omar Haddad mentioned you": (ActivityKind.MENTION, "Omar Haddad"),
        "Jane Doe replied": (ActivityKind.REPLY, "Jane Doe"),
        "Alex Example reacted to your message": (ActivityKind.REACTION, "Alex Example"),
        "Jane Doe invited you to Project Atlas review": (ActivityKind.INVITE, "Jane Doe"),
        "Omar Haddad forwarded you Jane Doe's meeting": (ActivityKind.MEETING, "Omar Haddad"),
        "Omar Haddad updated Project Atlas sync": (ActivityKind.MEETING, "Omar Haddad"),
        "Missed call from Alex Example": (ActivityKind.MISSED_CALL, "Alex Example"),
        "Weekly digest is ready": (ActivityKind.OTHER, ""),
    }

    assert {title: classify(title) for title in cases} == cases


def test_feed_items_are_parsed_in_order_with_their_fields() -> None:
    root = teams_screen(TeamsApp.ACTIVITY, activity_feed([MENTION, REPLY, CALL]))

    items = parse_activity(root)

    assert [item.to_record() for item in items] == [
        {
            "kind": "mention",
            "person": "Omar Haddad",
            "title": "Omar Haddad mentioned you",
            "preview": "can you check the budget",
            "time_label": "2:05 PM",
            "location": "Project Atlas",
        },
        {
            "kind": "reply",
            "person": "Jane Doe",
            "title": "Jane Doe replied",
            "preview": "done",
            "time_label": "11/12",
            "location": "General",
        },
        {
            "kind": "missed_call",
            "person": "Alex Example",
            "title": "Missed call from Alex Example",
            "preview": "",
            "time_label": "Yesterday",
            "location": "",
        },
    ]


def test_unread_filter_is_switched_on_for_the_read_then_back_and_the_rail_restored() -> None:
    gateway = FakeTeamsGateway([])
    desk = ActivityDesk(gateway)

    items = reading(gateway).activity(ActivityFilter.UNREAD, 30)

    assert [item.title for item in items] == ["Omar Haddad mentioned you"]
    assert gateway.pressed == [
        rail_handle(TeamsApp.ACTIVITY),
        "toggle:Unread",
        "toggle:Unread",
        rail_handle(TeamsApp.CHAT),
    ]
    assert desk.app is TeamsApp.CHAT and desk.unread is False


def test_all_switches_an_owner_filter_off_and_puts_it_back() -> None:
    gateway = FakeTeamsGateway([])
    desk = ActivityDesk(gateway, unread=True)

    items = reading(gateway).activity(ActivityFilter.ALL, 2)

    assert [item.kind for item in items] == [ActivityKind.MENTION, ActivityKind.REPLY]
    assert gateway.pressed.count("toggle:Unread") == 2
    assert desk.unread is True
    assert gateway.pressed[-1] == rail_handle(TeamsApp.CHAT)


def test_mentions_filter_leaves_unread_alone() -> None:
    gateway = FakeTeamsGateway([])
    desk = ActivityDesk(gateway)

    items = reading(gateway).activity(ActivityFilter.MENTIONS, 30)

    assert [item.kind for item in items] == [ActivityKind.MENTION]
    assert "toggle:Unread" not in gateway.pressed
    assert gateway.pressed.count("toggle:@Mentions") == 2
    assert desk.mentions is False
