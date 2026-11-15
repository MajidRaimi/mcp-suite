from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any

from office_mcp.shared.node import Node


@dataclass(frozen=True)
class Channel:
    team: str
    team_id: str
    channel: str
    channel_id: str
    unread: bool
    private: bool
    channels_listed: bool

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ChannelRow:
    channel: Channel
    node: Node


@dataclass(frozen=True)
class ChannelReply:
    id: str
    author: str
    sent: str
    text: str

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ChannelPost:
    post_id: str
    author: str
    sent: str
    subject: str
    text: str
    reply_count: int
    replies: tuple[ChannelReply, ...]

    def to_record(self) -> dict[str, Any]:
        return {
            "post_id": self.post_id,
            "author": self.author,
            "sent": self.sent,
            "subject": self.subject,
            "text": self.text,
            "reply_count": self.reply_count,
            "replies": [reply.to_record() for reply in self.replies],
        }


@dataclass(frozen=True)
class ChannelPosts:
    team: str
    channel: str
    posts: tuple[ChannelPost, ...]

    def to_record(self) -> dict[str, Any]:
        return {
            "team": self.team,
            "channel": self.channel,
            "posts": [post.to_record() for post in self.posts],
        }
