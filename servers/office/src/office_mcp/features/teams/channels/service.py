from __future__ import annotations

from datetime import tzinfo

from office_mcp.features.teams.channels.models import ChannelPosts, ChannelRow
from office_mcp.features.teams.channels.parse import parse_channels
from office_mcp.features.teams.channels.posts import parse_posts, posts_shown
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.features.teams.reading.view import ReadScope
from office_mcp.shared.errors import AmbiguousTargetError, TargetNotFoundError, UiTimeoutError
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

TITLE_SEPARATOR = " | "
TITLED_PARTS = 3
MAX_CANDIDATES = 8


def name_matches(wanted: str, name: str, exact: bool) -> bool:
    target = normalize(wanted)
    folded = normalize(name)
    return folded == target if exact else bool(target) and target in folded


def filter_channels(
    rows: list[ChannelRow], team: str | None, unread_only: bool
) -> list[ChannelRow]:
    picked = rows
    if team:
        wanted = normalize(team)
        picked = [row for row in picked if wanted in normalize(row.channel.team)]
    if unread_only:
        picked = [row for row in picked if row.channel.unread]
    return picked


def resolve_channel(rows: list[ChannelRow], team: str, channel: str) -> ChannelRow:
    listed = [row for row in rows if row.channel.channels_listed]
    for exact in (True, False):
        found = [
            row
            for row in listed
            if name_matches(team, row.channel.team, exact)
            and name_matches(channel, row.channel.channel, exact)
        ]
        if len(found) == 1:
            return found[0]
        if len(found) > 1:
            names = [f"{row.channel.team} / {row.channel.channel}" for row in found]
            raise AmbiguousTargetError(
                f"{team!r} / {channel!r} matches several channels", names[:MAX_CANDIDATES]
            )
    raise TargetNotFoundError(not_found_message(rows, team, channel))


def not_found_message(rows: list[ChannelRow], team: str, channel: str) -> str:
    collapsed = [
        row.channel.team
        for row in rows
        if not row.channel.channels_listed and name_matches(team, row.channel.team, exact=False)
    ]
    if collapsed:
        return (
            f"the team {collapsed[0]!r} is collapsed in Teams, so its channels are not listed; "
            "expand it once in Teams and try again"
        )
    return f"no channel {channel!r} in a team named {team!r} is listed in Teams"


def shows_channel(root: Node, row: ChannelRow) -> bool:
    if not posts_shown(root):
        return False
    parts = [normalize(part) for part in root.title.split(TITLE_SEPARATOR)]
    if len(parts) >= TITLED_PARTS:
        team, channel = parts[-2:]
        return team == normalize(row.channel.team) and channel == normalize(row.channel.channel)
    current = next(
        (
            entry
            for entry in parse_channels(root)
            if entry.channel.channel_id == row.channel.channel_id
        ),
        None,
    )
    return current is not None and current.node.selected is True


class ChannelReader:
    def __init__(self, gateway: TeamsGateway, waiter: Waiter, zone: tzinfo) -> None:
        self._gateway = gateway
        self._waiter = waiter
        self._zone = zone

    def channels(self, scope: ReadScope, team: str | None, unread_only: bool) -> list[ChannelRow]:
        return filter_channels(self._listed(scope), team, unread_only)

    def read(
        self, scope: ReadScope, team: str, channel: str, limit: int, replies: bool
    ) -> ChannelPosts:
        row = resolve_channel(self._listed(scope), team, channel)
        root = self._open(row)
        posts = parse_posts(root, self._zone, replies)
        shown = posts[-limit:] if limit > 0 else posts
        return ChannelPosts(team=row.channel.team, channel=row.channel.channel, posts=tuple(shown))

    def _listed(self, scope: ReadScope) -> list[ChannelRow]:
        scope.show_app(TeamsApp.TEAMS)
        return self._waiter.require(
            lambda: parse_channels(self._gateway.snapshot()) or None,
            lambda: UiTimeoutError("the Teams list of teams and channels did not load in time"),
        )

    def _open(self, row: ChannelRow) -> Node:
        root = self._gateway.snapshot()
        if not shows_channel(root, row):
            self._gateway.press(row.node)
            root = self._waiter.require(
                lambda: self._showing(row),
                lambda: UiTimeoutError(
                    f"Teams did not open the channel {row.channel.channel!r} in time"
                ),
            )
        return self._settled(root)

    def _showing(self, row: ChannelRow) -> Node | None:
        root = self._gateway.snapshot()
        return root if shows_channel(root, row) else None

    def _settled(self, root: Node) -> Node:
        latest = [root]

        def sample() -> int:
            latest[0] = self._gateway.snapshot()
            return sum(1 for _ in latest[0].walk())

        self._waiter.settle(sample)
        return latest[0]
