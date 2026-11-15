from __future__ import annotations

import re

from office_mcp.features.calendar.models import MyResponse, Peek, Responses
from office_mcp.features.calendar.parse import clean
from office_mcp.shared.node import STATIC_TEXT, Node
from office_mcp.shared.text import normalize

CALLOUT_CLASS = "ms-Callout"
DIALOG = "AXApplicationDialog"
VIEW_EVENT = "view event"
GROUP_ROLES = frozenset({"AXGroup", "AXList", "AXSection"})
BUTTON_ROLES = frozenset({"AXButton", "AXMenuButton", "AXPopUpButton"})
ORGANIZER_CARD_PREFIX = "Opens card for "
INVITED_MARK = "invited you"
RESPOND_FORM = "respond form"
EMAIL_ORGANIZER = "email organizer"
MESSAGE_AREA = "add a message"
MESSAGE_PLACEHOLDER = "Add a message (optional)"
ACCEPT_ID = "Accept"
DECLINE_ID = "Decline"
MORE_ID = "calendar_ellipses_menu"
TENTATIVE_ITEM = "Tentative"
JOIN = "join"
CHECKED = "1"
YOU_PREFIX = "you "
SERIES_MARKS = ("series", "occurrence", "this event", "all events")
STATE_LABELS: dict[str, MyResponse] = {
    "accepted": MyResponse.ACCEPTED,
    "you accepted": MyResponse.ACCEPTED,
    "tentatively accepted": MyResponse.TENTATIVE,
    "you tentatively accepted": MyResponse.TENTATIVE,
    "declined": MyResponse.DECLINED,
    "you declined": MyResponse.DECLINED,
}
COUNT_PATTERN = re.compile(
    r"(?P<kind>tentatively accepted|accepted|declined|tentative|didn't respond|did not respond|"
    r"no response|not responded)\s+(?P<count>\d+)",
    re.IGNORECASE,
)
COUNT_FIELDS: dict[str, str] = {
    "accepted": "accepted",
    "tentatively accepted": "tentative",
    "tentative": "tentative",
    "declined": "declined",
    "didn't respond": "not_responded",
    "did not respond": "not_responded",
    "no response": "not_responded",
    "not responded": "not_responded",
}


def is_callout(node: Node) -> bool:
    return any(name.startswith(CALLOUT_CLASS) for name in node.classes)


def find_peek(root: Node) -> Node | None:
    callout = root.find(is_callout)
    if callout is not None:
        return callout
    return root.find(
        lambda node: (
            (node.subrole == DIALOG or node.role == DIALOG)
            and node.find(lambda child: normalize(child.label()) == VIEW_EVENT) is not None
        )
    )


def statics(node: Node) -> list[str]:
    return [clean(item.value) for item in node.walk() if item.role == STATIC_TEXT and item.value]


def section(peek: Node, name: str) -> Node | None:
    wanted = normalize(name)
    return peek.find(lambda node: node.role in GROUP_ROLES and normalize(node.label()) == wanted)


def section_lines(peek: Node, name: str) -> list[str]:
    found = section(peek, name)
    if found is None:
        return []
    return [line for line in statics(found) if normalize(line) != normalize(name)]


def time_lines(peek: Node) -> tuple[str, str | None]:
    lines = section_lines(peek, "Time")
    timed = [line for line in lines if re.search(r"\d{1,2}:\d{2}", line)]
    others = [line for line in lines if line not in timed]
    return (timed[0] if timed else ""), (" ".join(others) or None)


def organizer(peek: Node) -> str:
    for node in peek.walk():
        label = clean(node.label())
        if node.role in BUTTON_ROLES and label.startswith(ORGANIZER_CARD_PREFIX):
            return label[len(ORGANIZER_CARD_PREFIX) :].strip()
    return ""


def invited_text(peek: Node) -> str:
    return next((line for line in statics(peek) if INVITED_MARK in line.casefold()), "")


def responses(peek: Node) -> Responses:
    counts: dict[str, int] = {}
    for line in statics(peek):
        for match in COUNT_PATTERN.finditer(line):
            field = COUNT_FIELDS[match["kind"].casefold()]
            counts[field] = counts.get(field, 0) + int(match["count"])
    return Responses(**counts)


def has_join_button(peek: Node) -> bool:
    return peek.find(lambda node: node.role in BUTTON_ROLES and is_join(node)) is not None


def is_join(node: Node) -> bool:
    return normalize(node.label()) == JOIN


def respond_form(peek: Node) -> Node | None:
    named = peek.find(lambda node: normalize(node.label()) == RESPOND_FORM)
    if named is not None:
        return named
    return peek if email_checkbox(peek) is not None else None


def email_checkbox(peek: Node) -> Node | None:
    return peek.find(
        lambda node: node.role == "AXCheckBox" and normalize(node.label()) == EMAIL_ORGANIZER
    )


def message_area(peek: Node) -> Node | None:
    return peek.find(
        lambda node: (
            node.role in ("AXTextArea", "AXTextField")
            and normalize(node.label()).startswith(MESSAGE_AREA)
        )
    )


def by_dom_id(peek: Node, dom_id: str) -> Node | None:
    return peek.find(lambda node: node.dom_id == dom_id)


def checkbox_state(node: Node) -> bool | None:
    if node.value in ("0", "1"):
        return node.value == CHECKED
    return node.selected


def my_response(peek: Node) -> MyResponse:
    for node in peek.walk():
        state = labelled_state(node)
        if state is not None:
            return state
    pressed = pressed_response(peek)
    if pressed is not None:
        return pressed
    if respond_form(peek) is None:
        return MyResponse.UNKNOWN if invited_text(peek) else MyResponse.ORGANIZER
    accept = by_dom_id(peek, ACCEPT_ID)
    if accept is not None and normalize(accept.label()) == "accept":
        return MyResponse.NOT_RESPONDED
    return MyResponse.UNKNOWN


def labelled_state(node: Node) -> MyResponse | None:
    label = normalize(node.label())
    if node.role in BUTTON_ROLES or (node.role == STATIC_TEXT and label.startswith(YOU_PREFIX)):
        return STATE_LABELS.get(label)
    return None


def pressed_response(peek: Node) -> MyResponse | None:
    for dom_id, state in ((ACCEPT_ID, MyResponse.ACCEPTED), (DECLINE_ID, MyResponse.DECLINED)):
        node = by_dom_id(peek, dom_id)
        if node is not None and node.value == CHECKED:
            return state
    return None


def series_prompt(root: Node) -> Node | None:
    for menu in (node for node in root.walk() if node.role == "AXMenu"):
        labels = [normalize(node.label()) for node in menu.walk() if node is not menu]
        if any(mark in label for label in labels for mark in SERIES_MARKS):
            return menu
    return None


def parse_peek(peek: Node) -> Peek:
    time_text, recurrence = time_lines(peek)
    title_lines = [line for line in section_lines(peek, "Title") if normalize(line) != VIEW_EVENT]
    return Peek(
        subject=title_lines[0] if title_lines else "",
        time_text=time_text,
        recurrence=recurrence,
        location=" ".join(section_lines(peek, "Location")),
        organizer=organizer(peek),
        invited_text=invited_text(peek),
        my_response=my_response(peek),
        responses=responses(peek),
        notes="\n".join(section_lines(peek, "Notes")),
        has_join_button=has_join_button(peek),
    )
