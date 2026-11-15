from __future__ import annotations

import re
from dataclasses import dataclass, field

from office_mcp.features.teams.channels.models import Channel, ChannelRow
from office_mcp.features.teams.chats.parse import own_find
from office_mcp.shared.node import Node

ROW_ROLE = "AXRow"
TEAM_PREFIX = "title-team-list-item-"
CHANNEL_PREFIX = "title-channel-list-item-"
SECTION_PREFIX = "Teams and channels "
TEAM_MARK = "Team "
CHANNEL_MARK = "Channel "
UNREAD = re.compile(r"\bUnread\b")
PRIVATE = re.compile(r"\bPrivate\b")


@dataclass
class TeamGroup:
    name: str
    team_id: str
    unread: bool
    node: Node
    channels: list[ChannelRow] = field(default_factory=list)


def parse_channels(root: Node) -> list[ChannelRow]:
    groups: list[TeamGroup] = []
    for row in root.walk():
        if row.role != ROW_ROLE:
            continue
        team_title = own_find(row, lambda node: node.dom_id.startswith(TEAM_PREFIX))
        if team_title is not None:
            groups.append(team_group(row, team_title))
            continue
        channel_title = own_find(row, lambda node: node.dom_id.startswith(CHANNEL_PREFIX))
        if channel_title is not None and groups:
            groups[-1].channels.append(channel_row(row, channel_title, groups[-1]))
    return [entry for group in groups for entry in group_entries(group)]


def team_group(row: Node, title_node: Node) -> TeamGroup:
    name = title_node.text() or title_node.label()
    flags = flags_text(row.label(), TEAM_MARK + name, name)
    return TeamGroup(
        name=name,
        team_id=title_node.dom_id.removeprefix(TEAM_PREFIX),
        unread=UNREAD.search(flags) is not None,
        node=row,
    )


def channel_row(row: Node, title_node: Node, group: TeamGroup) -> ChannelRow:
    name = title_node.text() or title_node.label()
    flags = flags_text(row.label(), CHANNEL_MARK + name, name).replace(group.name, " ")
    channel = Channel(
        team=group.name,
        team_id=group.team_id,
        channel=name,
        channel_id=title_node.dom_id.removeprefix(CHANNEL_PREFIX),
        unread=UNREAD.search(flags) is not None,
        private=PRIVATE.search(flags) is not None,
        channels_listed=True,
    )
    return ChannelRow(channel=channel, node=row)


def flags_text(title: str, marked_name: str, name: str) -> str:
    stripped = title.removeprefix(SECTION_PREFIX)
    index = stripped.find(marked_name)
    if index >= 0:
        return stripped[:index]
    return stripped.replace(name, " ") if name else stripped


def group_entries(group: TeamGroup) -> list[ChannelRow]:
    if group.channels:
        return group.channels
    collapsed = Channel(
        team=group.name,
        team_id=group.team_id,
        channel="",
        channel_id="",
        unread=group.unread,
        private=False,
        channels_listed=False,
    )
    return [ChannelRow(channel=collapsed, node=group.node)]
