from __future__ import annotations

from collections.abc import Callable
from datetime import date

from office_mcp.shared.node import Node

TODAY = date(2026, 10, 20)


def text(value: str) -> Node:
    return Node(role="AXStaticText", value=value)


def chat_row(
    thread: str,
    name: str,
    preview: str,
    time_label: str,
    kind: str = "Chat",
    presence: str = "Available",
    unread: bool = False,
    muted: bool = False,
) -> Node:
    parts = ["Unread message"] if unread else []
    parts += [kind, name]
    if kind == "Chat" and presence:
        parts.append(presence)
    parts += ["Last message", preview]
    if muted:
        parts.append("Muted")
    parts.append(time_label)
    children = [
        Node(role="AXGroup", dom_id=f"title-chat-list-item_{thread}", children=(text(name),)),
        Node(role="AXGroup", dom_id=f"time-chat-list-item_{thread}", children=(text(time_label),)),
        Node(
            role="AXGroup",
            dom_id=f"message-preview-chat-list-item_{thread}",
            children=(text(preview),),
        ),
    ]
    if unread:
        children.insert(0, Node(role="AXGroup", dom_id="unread-label", children=(text("Unread"),)))
    return Node(role="AXRow", title=" ".join(parts), children=tuple(children), handle=thread)


def section_row(title: str, rows: list[Node] | None = None) -> Node:
    return Node(role="AXRow", title=title, children=tuple(rows or []), handle=title)


def message(
    message_id: str,
    author: str,
    body: str,
    stamp: str,
    mine: bool = False,
    show_author: bool = True,
) -> list[Node]:
    prefix = "fui-ChatMyMessage" if mine else "fui-ChatMessage"
    header: list[Node] = [
        Node(
            role="AXGroup",
            description=f"{stamp}.",
            dom_id=f"timestamp-{message_id}",
            classes=(f"{prefix}__timestamp",),
            children=(text("short"),),
        )
    ]
    if show_author:
        header.append(Node(role="AXGroup", dom_id=f"author-{message_id}", children=(text(author),)))
    return [
        Node(role="AXHeading", title=f"{body} by {author}", value="4"),
        Node(role="AXGroup", children=tuple(header)),
        Node(
            role="AXGroup",
            title=f"{body} {author} {stamp}.",
            dom_id=f"message-body-{message_id}",
            classes=(f"{prefix}__body",),
            children=(
                Node(role="AXButton", title="More message options"),
                Node(
                    role="AXGroup",
                    title=body,
                    dom_id=f"content-{message_id}",
                    children=(text(body),),
                ),
            ),
        ),
    ]


def window(title: str, rows: list[Node], messages: list[Node]) -> Node:
    pane = Node(role="AXGroup", dom_id="chat-pane-list", children=tuple(messages))
    return Node(
        role="AXWindow",
        title=title,
        children=(Node(role="AXGroup", children=tuple(rows)), pane),
    )


class FakeGateway:
    def __init__(self, screens: list[Node], on_press: Callable[[Node], None] | None = None) -> None:
        self.screens = screens
        self.pressed: list[object] = []
        self.scrolled: list[str] = []
        self._on_press = on_press

    def snapshot(self) -> Node:
        return self.screens[0] if len(self.screens) == 1 else self.screens.pop(0)

    def press(self, node: Node) -> None:
        self.pressed.append(node.handle)
        if self._on_press:
            self._on_press(node)

    def scroll_into_view(self, node: Node) -> None:
        self.scrolled.append(node.dom_id)


class FakeClock:
    def __init__(self) -> None:
        self.now = 0.0

    def __call__(self) -> float:
        return self.now

    def sleep(self, seconds: float) -> None:
        self.now += seconds
