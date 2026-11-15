from __future__ import annotations

from collections.abc import Callable

from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import chat_header, static
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.node import Node

RAIL_TITLES = {
    TeamsApp.ACTIVITY: "Activity",
    TeamsApp.CHAT: "Chat",
    TeamsApp.TEAMS: "Teams",
    TeamsApp.CALENDAR: "Calendar",
}


class ScriptedTeamsGateway(FakeTeamsGateway):
    def __init__(
        self, screens: list[Node], on_value: Callable[[Node, str], None] | None = None
    ) -> None:
        super().__init__(screens)
        self.on_value = on_value

    def set_value(self, node: Node, text: str) -> None:
        super().set_value(node, text)
        if self.on_value is not None:
            self.on_value(node, text)


def rail_handle(app: TeamsApp) -> str:
    return f"rail:{RAIL_TITLES[app]}"


def full_rail(selected: TeamsApp | None) -> Node:
    toggles = tuple(
        Node(
            role="AXCheckBox",
            subrole="AXToggleButton",
            title=title,
            dom_id=app.value,
            value="1" if app is selected else "0",
            handle=rail_handle(app),
        )
        for app, title in RAIL_TITLES.items()
    )
    return Node(role="AXGroup", subrole="AXLandmarkNavigation", title="Apps", children=toggles)


def teams_screen(
    selected: TeamsApp | None,
    *content: Node,
    title: str = "Microsoft Teams",
    thread: str | None = None,
) -> Node:
    header = (chat_header(thread),) if thread is not None else ()
    return Node(role="AXWindow", title=title, children=(full_rail(selected), *header, *content))


def labeled(dom_id: str, text: str) -> Node:
    return Node(role="AXStaticText", value=text, dom_id=dom_id)


def activity_item(key: str, title: str, preview: str, time_label: str, location: str) -> Node:
    return Node(
        role="AXStaticText",
        value=f"{title} {preview} {time_label} {location}",
        classes=("ui-tree__title",),
        children=(
            labeled(f"activity-feed-item-title-{key}", title),
            labeled(f"activity-feed-item-message-preview-{key}", preview),
            labeled(f"activity-feed-item-timestamp-{key}", time_label),
            labeled(f"activity-feed-item-location-{key}", location),
        ),
    )


def filter_toggle(label: str, on: bool) -> Node:
    return Node(
        role="AXCheckBox",
        subrole="AXToggleButton",
        title=label,
        value="1" if on else "0",
        handle=f"toggle:{label.split()[0]}",
    )


def activity_feed(items: list[Node], unread: bool = False, mentions: bool = False) -> Node:
    toolbar = Node(
        role="AXToolbar",
        dom_id="upfront-filter-toolbar-id",
        children=(
            filter_toggle("Unread ⌥⌘U", unread),
            filter_toggle("@Mentions ⌥⌘B", mentions),
            filter_toggle("Tag mentions", False),
        ),
    )
    feed = Node(role="AXList", title="Activity", children=tuple(items))
    heading = Node(role="AXHeading", dom_id="feed-header", title="Feed")
    return Node(
        role="AXGroup",
        subrole="AXLandmarkRegion",
        title="Activity",
        children=(heading, toolbar, feed),
    )


def search_box(value: str = "") -> Node:
    return Node(role="AXComboBox", dom_id="ms-searchux-input", value=value, handle="search-box")


def search_popup(*texts: str) -> Node:
    return Node(
        role="AXList",
        title="Search suggestions",
        dom_id="ms-searchux-popup-0a1b",
        children=tuple(static(text) for text in texts),
    )


def team_row(team_id: str, name: str, unread: bool = False) -> Node:
    flag = "Unread " if unread else ""
    return Node(
        role="AXRow",
        subrole="AXOutlineRow",
        title=f"Teams and channels {flag}Team {name} More options",
        handle=f"team:{name}",
        children=(
            Node(
                role="AXGroup",
                dom_id=f"title-team-list-item-{team_id}",
                children=(static(name),),
            ),
        ),
    )


def channel_row(
    channel_id: str,
    team: str,
    name: str,
    unread: bool = False,
    private: bool = False,
    selected: bool = False,
) -> Node:
    flags = ("Unread " if unread else "") + ("Private " if private else "")
    return Node(
        role="AXRow",
        subrole="AXOutlineRow",
        title=f"Teams and channels {team} {flags}Channel {name} 10/19",
        handle=f"channel:{name}",
        selected=selected,
        children=(
            Node(
                role="AXGroup",
                dom_id=f"title-channel-list-item-{channel_id}",
                children=(static(name),),
            ),
        ),
    )


def see_all_row() -> Node:
    return Node(role="AXRow", title="See all channels", handle="see-all")


def teams_tree(*rows: Node) -> Node:
    return Node(role="AXOutline", title="Teams", children=rows)


def post_field(prefix: str, item_id: str, text: str, role: str = "AXGroup") -> Node:
    return Node(role=role, dom_id=f"{prefix}{item_id}", children=(static(text),))


def channel_reply(reply_id: str, author: str, stamp: str, text: str) -> Node:
    return Node(
        role="AXGroup",
        children=(
            post_field("author-", reply_id, author),
            Node(role="AXGroup", dom_id=f"timestamp-{reply_id}", title=f"{stamp}."),
            post_field("content-", reply_id, text),
        ),
    )


def channel_post(
    post_id: str,
    author: str,
    stamp: str,
    subject: str,
    text: str,
    replies: list[Node] | None = None,
    reply_title: str = "",
) -> list[Node]:
    summary = Node(
        role="AXGroup",
        subrole="AXApplicationGroup",
        dom_id=f"reply-chain-summary-{post_id}",
        children=(
            Node(
                role="AXGroup",
                dom_id=f"post-message-renderer-{post_id}",
                children=(
                    post_field("author-", post_id, author),
                    Node(role="AXGroup", dom_id=f"timestamp-{post_id}", title=f"{stamp}."),
                    post_field("subject-line-", post_id, subject, role="AXHeading"),
                    post_field("content-", post_id, text),
                ),
            ),
        ),
    )
    if replies is None:
        return [summary]
    surface = Node(
        role="AXGroup",
        title=reply_title or f"{len(replies)} replies",
        dom_id="response-surface-19:general@thread.tacv2",
        children=tuple(replies),
    )
    return [summary, surface]


def posts_pane(*posts: Node) -> Node:
    tab = Node(role="AXRadioButton", dom_id="com.microsoft.channeltabs.posts", title="Posts")
    return Node(role="AXGroup", title="Posts", children=(tab, *posts))


def quoted_content(
    message_id: str, quote_title: str, quote_pieces: list[str], reply: str, title: str = ""
) -> Node:
    quote = Node(
        role="AXGroup",
        subrole="AXApplicationGroup",
        title=quote_title,
        children=tuple(static(piece) for piece in quote_pieces),
    )
    return Node(
        role="AXGroup",
        title=title,
        dom_id=f"content-{message_id}",
        children=(quote, static(reply)),
    )


def message_with_content(
    message_id: str, author: str, content: Node, mine: bool = False, edited: bool = False
) -> list[Node]:
    prefix = "fui-ChatMyMessage" if mine else "fui-ChatMessage"
    header: list[Node] = [
        Node(role="AXGroup", dom_id=f"author-{message_id}", children=(static(author),)),
    ]
    if edited:
        header.append(Node(role="AXStaticText", value="Edited"))
    return [
        Node(role="AXHeading", title=f"preview by {author}", value="4"),
        Node(role="AXGroup", children=tuple(header)),
        Node(
            role="AXGroup",
            dom_id=f"message-body-{message_id}",
            classes=(f"{prefix}__body",),
            handle=f"body:{message_id}",
            children=(content,),
        ),
    ]


def conversation_post(
    post_id: str,
    author: str,
    stamp: str,
    subject: str,
    text: str,
    summary: str = "",
    replies: tuple[tuple[str, str, str, str], ...] = (),
) -> Node:
    def body(item_id: str, who: str, when: str, words: str, heading: str = "") -> Node:
        content = (
            Node(
                role="AXGroup",
                dom_id=f"content-{item_id}",
                children=(
                    *(
                        (post_field("subject-line-", item_id, heading, role="AXHeading"),)
                        if heading
                        else ()
                    ),
                    static(words),
                ),
            ),
        )
        return Node(
            role="AXGroup",
            subrole="AXApplicationGroup",
            dom_id=f"message-body-{item_id}",
            children=(
                post_field("author-", item_id, who),
                Node(role="AXGroup", dom_id=f"timestamp-{item_id}", title=f"{when}."),
                *content,
            ),
        )

    kids = [body(post_id, author, stamp, text, subject)]
    if summary:
        kids.append(
            Node(
                role="AXGroup",
                subrole="AXApplicationGroup",
                title=summary,
                children=(
                    Node(role="AXButton", dom_id=f"rc-sum-cont-btn-{post_id}", title=summary),
                ),
            )
        )
    kids.extend(body(*reply) for reply in replies)
    return Node(role="AXGroup", dom_id=f"menur{post_id[-3:]}", children=tuple(kids))


def conversation_pane(*threads: Node) -> Node:
    tab = Node(
        role="AXRadioButton",
        subrole="AXTabButton",
        dom_id="com.microsoft.channeltabs.conversation",
        title="Conversation",
        value="1",
    )
    pane = Node(role="AXGroup", dom_id="chat-pane-list", children=threads)
    return Node(role="AXGroup", children=(tab, pane))
