from __future__ import annotations

import re
from collections.abc import Iterator
from datetime import tzinfo

from office_mcp.features.teams.channels.models import ChannelPost, ChannelReply
from office_mcp.shared.clock import message_time
from office_mcp.shared.node import Node

SUMMARY_PREFIX = "reply-chain-summary-"
REPLIES_PREFIX = "response-surface-"
AUTHOR_PREFIX = "author-"
TIME_PREFIX = "timestamp-"
SUBJECT_PREFIX = "subject-line-"
CONTENT_PREFIX = "content-"
POSTS_TAB_ID = "com.microsoft.channeltabs.posts"
CONVERSATION_TAB_ID = "com.microsoft.channeltabs.conversation"
POST_TAB_IDS = frozenset({POSTS_TAB_ID, CONVERSATION_TAB_ID})
PANE_ID = "chat-pane-list"
BODY_PREFIX = "message-body-"
REPLY_SUMMARY_PREFIX = "rc-sum-cont-btn-"
REPLY_COUNT = re.compile(r"(?P<count>\d+)\s+repl", re.IGNORECASE)


def is_summary(node: Node) -> bool:
    return node.dom_id.startswith(SUMMARY_PREFIX)


def is_replies(node: Node) -> bool:
    return node.dom_id.startswith(REPLIES_PREFIX)


def is_body(node: Node) -> bool:
    return node.dom_id.startswith(BODY_PREFIX)


def posts_shown(root: Node) -> bool:
    return root.find(lambda node: node.dom_id in POST_TAB_IDS or is_summary(node)) is not None


def thread_parts(node: Node) -> Iterator[Node]:
    if is_summary(node) or is_replies(node):
        yield node
        return
    for child in node.children:
        yield from thread_parts(child)


def walk_outside_replies(node: Node) -> Iterator[Node]:
    yield node
    for child in node.children:
        if not is_replies(child):
            yield from walk_outside_replies(child)


def parse_posts(root: Node, zone: tzinfo, with_replies: bool) -> list[ChannelPost]:
    if root.find(is_summary) is None:
        return parse_conversation(root, zone, with_replies)
    posts: list[ChannelPost] = []
    for part in thread_parts(root):
        if is_summary(part):
            posts.append(build_post(part, zone, with_replies))
        elif posts:
            posts[-1] = attach_replies(posts[-1], part, zone, with_replies)
    return posts


def build_post(summary: Node, zone: tzinfo, with_replies: bool) -> ChannelPost:
    post_id = summary.dom_id.removeprefix(SUMMARY_PREFIX)
    by_id = {node.dom_id: node for node in walk_outside_replies(summary) if node.dom_id}
    post = ChannelPost(
        post_id=post_id,
        author=field_text(by_id, AUTHOR_PREFIX, post_id),
        sent=sent_text(post_id, field_node(by_id, TIME_PREFIX, post_id), zone),
        subject=field_text(by_id, SUBJECT_PREFIX, post_id),
        text=field_text(by_id, CONTENT_PREFIX, post_id),
        reply_count=0,
        replies=(),
    )
    nested = summary.find(is_replies)
    return attach_replies(post, nested, zone, with_replies) if nested is not None else post


def field_node(by_id: dict[str, Node], prefix: str, item_id: str) -> Node | None:
    exact = by_id.get(prefix + item_id)
    if exact is not None:
        return exact
    return next((node for key, node in by_id.items() if key.startswith(prefix)), None)


def field_text(by_id: dict[str, Node], prefix: str, item_id: str) -> str:
    node = field_node(by_id, prefix, item_id)
    return node.label() if node is not None else ""


def sent_text(item_id: str, stamp: Node | None, zone: tzinfo) -> str:
    sent = message_time(item_id, zone)
    if sent is not None:
        return sent.isoformat(timespec="seconds")
    return stamp.label().rstrip(".") if stamp is not None else ""


def attach_replies(
    post: ChannelPost, surface: Node, zone: tzinfo, with_replies: bool
) -> ChannelPost:
    replies = parse_replies(surface, zone)
    counted = REPLY_COUNT.search(surface.title or surface.description)
    count = int(counted["count"]) if counted else len(replies)
    return ChannelPost(
        post_id=post.post_id,
        author=post.author,
        sent=post.sent,
        subject=post.subject,
        text=post.text,
        reply_count=post.reply_count + count,
        replies=post.replies + (replies if with_replies else ()),
    )


def parse_replies(surface: Node, zone: tzinfo) -> tuple[ChannelReply, ...]:
    by_id = {node.dom_id: node for node in surface.walk() if node.dom_id}
    ids = [key.removeprefix(CONTENT_PREFIX) for key in by_id if key.startswith(CONTENT_PREFIX)]
    return tuple(
        ChannelReply(
            id=reply_id,
            author=exact_text(by_id, AUTHOR_PREFIX + reply_id),
            sent=sent_text(reply_id, by_id.get(TIME_PREFIX + reply_id), zone),
            text=exact_text(by_id, CONTENT_PREFIX + reply_id),
        )
        for reply_id in ids
    )


def exact_text(by_id: dict[str, Node], dom_id: str) -> str:
    node = by_id.get(dom_id)
    return node.label() if node is not None else ""


def parse_conversation(root: Node, zone: tzinfo, with_replies: bool) -> list[ChannelPost]:
    pane = root.find(lambda node: node.dom_id == PANE_ID)
    if pane is None:
        return []
    posts = (thread_post(thread, zone, with_replies) for thread in pane.children)
    return [post for post in posts if post is not None]


def thread_post(thread: Node, zone: tzinfo, with_replies: bool) -> ChannelPost | None:
    bodies = [node for node in thread.walk() if is_body(node)]
    if not bodies:
        return None
    post_id = bodies[0].dom_id.removeprefix(BODY_PREFIX)
    by_id = {node.dom_id: node for node in thread.walk() if node.dom_id}
    replies = tuple(
        conversation_reply(body.dom_id.removeprefix(BODY_PREFIX), by_id, zone)
        for body in bodies[1:]
    )
    return ChannelPost(
        post_id=post_id,
        author=exact_text(by_id, AUTHOR_PREFIX + post_id),
        sent=sent_text(post_id, by_id.get(TIME_PREFIX + post_id), zone),
        subject=exact_text(by_id, SUBJECT_PREFIX + post_id),
        text=exact_text(by_id, CONTENT_PREFIX + post_id),
        reply_count=summary_count(thread) or len(replies),
        replies=replies if with_replies else (),
    )


def summary_count(thread: Node) -> int:
    summary = thread.find(lambda node: node.dom_id.startswith(REPLY_SUMMARY_PREFIX))
    counted = REPLY_COUNT.search(summary.label()) if summary is not None else None
    return int(counted["count"]) if counted else 0


def conversation_reply(reply_id: str, by_id: dict[str, Node], zone: tzinfo) -> ChannelReply:
    return ChannelReply(
        id=reply_id,
        author=exact_text(by_id, AUTHOR_PREFIX + reply_id),
        sent=sent_text(reply_id, by_id.get(TIME_PREFIX + reply_id), zone),
        text=exact_text(by_id, CONTENT_PREFIX + reply_id),
    )
