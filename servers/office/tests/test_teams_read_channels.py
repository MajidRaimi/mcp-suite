from __future__ import annotations

import pytest
from fakes.drivers import fake_runtime, teams_rig
from fakes.p6_trees import (
    channel_post,
    channel_reply,
    channel_row,
    conversation_pane,
    conversation_post,
    posts_pane,
    rail_handle,
    see_all_row,
    team_row,
    teams_screen,
    teams_tree,
)
from fakes.teams_gateway import FakeTeamsGateway

from office_mcp.features.teams.channels.models import Channel
from office_mcp.features.teams.channels.parse import parse_channels
from office_mcp.features.teams.channels.posts import parse_posts
from office_mcp.features.teams.channels.service import shows_channel
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.features.teams.reading.factory import build_teams_reading
from office_mcp.features.teams.reading.service import TeamsReadingService
from office_mcp.shared.errors import AmbiguousTargetError, TargetNotFoundError
from office_mcp.shared.node import Node

ATLAS = "19:atlas@thread.tacv2"
GENERAL = "19:general@thread.tacv2"
BUDGET = "19:budget@thread.tacv2"
CREW = "19:crew@thread.tacv2"
OMAR_THREAD = "19:omar@unq.gbl.spaces"
POST = "1790000100000"
REPLY_ONE = "1790000200000"
REPLY_TWO = "1790000300000"


def tree(open_channel: str) -> Node:
    return teams_tree(
        team_row(ATLAS, "Project Atlas", unread=True),
        channel_row(GENERAL, "Project Atlas", "General", selected=open_channel == "General"),
        channel_row(
            BUDGET,
            "Project Atlas",
            "Budget",
            unread=True,
            private=True,
            selected=open_channel == "Budget",
        ),
        see_all_row(),
        team_row(CREW, "Design Crew"),
    )


def budget_posts() -> list[Node]:
    return [
        *channel_post(
            POST,
            "Jane Doe",
            "October 19, 2026 9:00 AM",
            "Q4 numbers",
            "Draft budget is ready",
            replies=[
                channel_reply(REPLY_ONE, "Omar Haddad", "October 19, 2026 9:10 AM", "Looks good"),
                channel_reply(REPLY_TWO, "Alex Example", "October 19, 2026 9:20 AM", "Approved"),
            ],
            reply_title="5 replies",
        ),
        *channel_post("card-7", "Planner", "October 20, 2026 8:00 AM", "", "Task assigned"),
    ]


class ChannelsDesk:
    def __init__(self, gateway: FakeTeamsGateway) -> None:
        self.gateway = gateway
        self.app = TeamsApp.CHAT
        self.channel = "General"
        gateway.on_press = self.press
        gateway.show(self.render())

    def render(self) -> Node:
        if self.app is TeamsApp.CHAT:
            return teams_screen(TeamsApp.CHAT, title="Chat | Omar Haddad", thread=OMAR_THREAD)
        posts = budget_posts() if self.channel == "Budget" else []
        return teams_screen(
            TeamsApp.TEAMS,
            tree(self.channel),
            posts_pane(*posts),
            title=f"Teams and Channels | Project Atlas | {self.channel}",
        )

    def press(self, node: Node) -> None:
        if node.handle == rail_handle(TeamsApp.TEAMS):
            self.app = TeamsApp.TEAMS
        elif node.handle == rail_handle(TeamsApp.CHAT):
            self.app = TeamsApp.CHAT
        elif isinstance(node.handle, str) and node.handle.startswith("channel:"):
            self.channel = node.handle.removeprefix("channel:")
        self.gateway.show(self.render())


def reading(gateway: FakeTeamsGateway) -> TeamsReadingService:
    return build_teams_reading(fake_runtime(), teams_rig(gateway).driver)


def test_tree_rows_become_channels_and_collapsed_teams() -> None:
    channels = [
        row.channel for row in parse_channels(teams_screen(TeamsApp.TEAMS, tree("General")))
    ]

    assert channels == [
        Channel("Project Atlas", ATLAS, "General", GENERAL, False, False, True),
        Channel("Project Atlas", ATLAS, "Budget", BUDGET, True, True, True),
        Channel("Design Crew", CREW, "", "", False, False, False),
    ]
    assert set(channels[0].to_record()) == {
        "team",
        "team_id",
        "channel",
        "channel_id",
        "unread",
        "private",
        "channels_listed",
    }


def test_collapsed_team_reports_its_own_unread_state() -> None:
    rows = parse_channels(teams_tree(team_row(CREW, "Design Crew", unread=True)))

    assert rows[0].channel.unread is True
    assert rows[0].channel.channels_listed is False


def test_posts_carry_subject_text_replies_and_counts() -> None:
    root = posts_pane(*budget_posts())

    first, card = parse_posts(root, fake_runtime().settings.zone.info, with_replies=True)

    assert first.post_id == POST and first.author == "Jane Doe"
    assert first.subject == "Q4 numbers" and first.text == "Draft budget is ready"
    assert first.sent == "2026-09-21T14:15:00+00:00"
    assert first.reply_count == 5
    assert [(reply.author, reply.text) for reply in first.replies] == [
        ("Omar Haddad", "Looks good"),
        ("Alex Example", "Approved"),
    ]
    assert card.sent == "October 20, 2026 8:00 AM"
    assert card.reply_count == 0 and card.replies == ()


def test_posts_without_replies_keep_the_count() -> None:
    first, _ = parse_posts(posts_pane(*budget_posts()), fake_runtime().settings.zone.info, False)

    assert first.reply_count == 5
    assert first.replies == ()


def test_channels_lists_filters_and_restores_the_rail() -> None:
    gateway = FakeTeamsGateway([])
    desk = ChannelsDesk(gateway)

    channels = reading(gateway).channels(team="atlas", unread_only=True)

    assert [channel.channel for channel in channels] == ["Budget"]
    assert gateway.pressed == [rail_handle(TeamsApp.TEAMS), rail_handle(TeamsApp.CHAT)]
    assert desk.app is TeamsApp.CHAT


def test_channel_read_opens_the_channel_and_restores_the_rail() -> None:
    gateway = FakeTeamsGateway([])
    desk = ChannelsDesk(gateway)

    posts = reading(gateway).channel_read("project atlas", "budget", 1, replies=True)

    record = posts.to_record()
    assert record["team"] == "Project Atlas" and record["channel"] == "Budget"
    assert [post["post_id"] for post in record["posts"]] == ["card-7"]
    assert gateway.pressed == [
        rail_handle(TeamsApp.TEAMS),
        "channel:Budget",
        rail_handle(TeamsApp.CHAT),
    ]
    assert desk.app is TeamsApp.CHAT


def test_channel_read_does_not_press_an_already_open_channel() -> None:
    gateway = FakeTeamsGateway([])
    ChannelsDesk(gateway)

    posts = reading(gateway).channel_read("Project Atlas", "General", 20, replies=False)

    assert posts.posts == ()
    assert "channel:General" not in gateway.pressed


def test_channel_in_a_collapsed_team_names_the_fix() -> None:
    gateway = FakeTeamsGateway([])
    ChannelsDesk(gateway)

    with pytest.raises(TargetNotFoundError, match="collapsed"):
        reading(gateway).channel_read("Design Crew", "General", 20, replies=True)

    assert gateway.pressed[-1] == rail_handle(TeamsApp.CHAT)


def test_vague_channel_names_are_ambiguous() -> None:
    gateway = FakeTeamsGateway([])
    ChannelsDesk(gateway)

    with pytest.raises(AmbiguousTargetError, match="several channels"):
        reading(gateway).channel_read("atlas", "e", 20, replies=True)


def test_open_channel_with_the_same_name_in_another_team_is_not_trusted() -> None:
    rows = parse_channels(teams_screen(TeamsApp.TEAMS, tree("General")))
    general = rows[0]
    elsewhere = teams_screen(
        TeamsApp.TEAMS,
        tree("General"),
        posts_pane(),
        title="Teams and Channels | Design Crew | General",
    )
    here = teams_screen(
        TeamsApp.TEAMS,
        tree("General"),
        posts_pane(),
        title="Teams and Channels | Project Atlas | General",
    )

    assert shows_channel(elsewhere, general) is False
    assert shows_channel(here, general) is True


def conversation_threads() -> list[Node]:
    return [
        conversation_post(
            POST,
            "Jane Doe",
            "October 19, 2026 9:00 AM",
            "Q4 numbers",
            "Draft budget is ready",
            summary="2 replies October 19, 2026 9:20 AM",
        ),
        conversation_post(
            "1790000400000",
            "Omar Haddad",
            "October 20, 2026 8:00 AM",
            "",
            "Room is booked",
            replies=(("1790000500000", "Alex Example", "October 20, 2026 8:05 AM", "Thanks"),),
        ),
    ]


def test_conversation_layout_posts_read_subject_text_and_reply_counts() -> None:
    root = conversation_pane(*conversation_threads())

    first, second = parse_posts(root, fake_runtime().settings.zone.info, with_replies=True)

    assert first.post_id == POST and first.author == "Jane Doe"
    assert first.subject == "Q4 numbers"
    assert "Draft budget is ready" in first.text
    assert first.sent == "2026-09-21T14:15:00+00:00"
    assert first.reply_count == 2 and first.replies == ()
    assert second.author == "Omar Haddad" and second.subject == ""
    assert second.reply_count == 1
    assert [(reply.author, reply.text) for reply in second.replies] == [("Alex Example", "Thanks")]


def test_conversation_layout_counts_as_an_open_channel() -> None:
    root = teams_screen(
        TeamsApp.TEAMS,
        tree("General"),
        conversation_pane(*conversation_threads()),
        title="Teams and Channels | Project Atlas | General",
    )
    row = next(row for row in parse_channels(root) if row.channel.channel == "General")

    assert shows_channel(root, row)
