from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, date, timedelta

from fakes.drivers import NO_IDLE_WAIT
from fakes.events import RecordingEvents
from fakes.guard import RecordingGuard
from fakes.pasteboard import FakePasteboard
from fakes.trees import AppRig, app_rig
from fakes.workspace import FakeWorkspace
from office_mcp.features.calendar.ax import AxCalendar
from office_mcp.features.calendar.locations import EventLocations
from office_mcp.features.calendar.models import MyResponse
from office_mcp.features.calendar.service import CalendarDeps, CalendarService
from office_mcp.features.teams.chats.opener import ChatOpener
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.driver import TeamsDriver
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.ax import AccessMode
from office_mcp.shared.errors import ElementNotFoundError
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.node import Node
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.text_input import TextWriter

TODAY = date(2026, 11, 15)
CHAT_THREAD = "19:alice@unq.gbl.spaces"
WORK_WEEK = "Work week"
WEEK = "Week"
VIEWS = ("Day", WORK_WEEK, WEEK, "Month")
MESSAGE_PLACEHOLDER = "Add a message (optional)"
MONTH_NAMES = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)
WEEKDAY_NAMES = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
RAIL = (("Activity", TeamsApp.ACTIVITY), ("Chat", TeamsApp.CHAT), ("Calendar", TeamsApp.CALENDAR))
RESPONSE_LABELS = {
    MyResponse.ACCEPTED: "Accepted",
    MyResponse.DECLINED: "Declined",
    MyResponse.TENTATIVE: "Tentatively accepted",
}


def long_date(day: date) -> str:
    return f"{WEEKDAY_NAMES[day.weekday()]}, {MONTH_NAMES[day.month - 1]} {day.day}, {day.year}"


def week_start(day: date) -> date:
    return day - timedelta(days=(day.weekday() + 1) % 7)


def range_label(first: date, last: date) -> str:
    start = f"{MONTH_NAMES[first.month - 1]} {first.day}"
    if first.year != last.year:
        return f"{start}, {first.year} \u2013 {MONTH_NAMES[last.month - 1]} {last.day}, {last.year}"
    if first.month != last.month:
        return f"{start} \u2013 {MONTH_NAMES[last.month - 1]} {last.day}, {last.year}"
    return f"{start}\u2013{last.day}, {last.year}"


@dataclass
class DeskEvent:
    subject: str
    day: date
    hours: str = "10:00 to 10:30"
    location: str = "Microsoft Teams Meeting"
    organizer: str = "Jane Doe"
    show_as: str = "Busy"
    recurring: bool = False
    response: MyResponse = MyResponse.NOT_RESPONDED
    invited: bool = True
    notes: str = "Agenda: review the numbers."

    def title(self) -> str:
        parts = [self.subject, self.hours, long_date(self.day)]
        if self.location:
            parts.append(self.location)
        parts.extend([f"By {self.organizer}", self.show_as])
        if self.recurring:
            parts.append("Recurring event")
        return ", ".join(parts)


@dataclass
class CalendarDesk:
    events: list[DeskEvent]
    app: TeamsApp = TeamsApp.CHAT
    view: str = WORK_WEEK
    week_first: date = field(default_factory=lambda: week_start(TODAY))
    peek_for: int | None = None
    view_menu: bool = False
    more_menu: bool = False
    series_menu: bool = False
    email_checked: bool = True
    message: str = ""
    series_prompt: bool = False
    apply_response: bool = True
    close_on_commit: bool = True
    peek_ignores_popups: bool = True
    more_items: tuple[str, ...] = ("Tentative", "Follow")
    picker_delay: int = 1
    pressed: list[str] = field(default_factory=list)
    keys: list[Shortcut] = field(default_factory=list)
    focused: list[str] = field(default_factory=list)
    snapshots: int = 0
    picker_hidden: int = 0

    def shown_days(self) -> tuple[date, date]:
        if self.view == WORK_WEEK:
            return self.week_first + timedelta(days=1), self.week_first + timedelta(days=5)
        return self.week_first, self.week_first + timedelta(days=6)

    def snapshot(self) -> Node:
        self.snapshots += 1
        return self.render()

    def render(self) -> Node:
        kids: list[Node] = [self.rail()]
        if self.app is TeamsApp.CHAT:
            kids.append(Node(role="AXGroup", dom_id=f"chat-header-{CHAT_THREAD}"))
        if self.app is TeamsApp.CALENDAR:
            kids.extend(self.calendar())
        return Node(role="AXWindow", title="Teams", children=tuple(kids))

    def rail(self) -> Node:
        toggles = tuple(
            Node(
                role="AXCheckBox",
                subrole="AXToggleButton",
                title=title,
                dom_id=app.value,
                value="1" if app is self.app else "0",
                handle=f"rail:{title}",
            )
            for title, app in RAIL
        )
        return Node(role="AXGroup", subrole="AXLandmarkNavigation", title="Apps", children=toggles)

    def calendar(self) -> list[Node]:
        first, last = self.shown_days()
        week = timedelta(days=7)
        picker = Node(
            role="AXPopUpButton",
            title=f"Change the view of your calendar, button. {self.view} selected",
            handle="picker",
        )
        rendered = self.picker_hidden <= 0
        self.picker_hidden -= 1
        toolbar = Node(
            role="AXToolbar",
            children=(
                Node(role="AXButton", title=f"Go to today {long_date(TODAY)}", handle="today"),
                Node(
                    role="AXButton",
                    title=f"Go to previous week {range_label(first - week, last - week)}",
                    handle="prev",
                ),
                Node(
                    role="AXButton",
                    title=f"Go to next week {range_label(first + week, last + week)}",
                    handle="next",
                ),
                Node(
                    role="AXButton",
                    title=f"{range_label(first, last)}, Jump to a specific date or date range.",
                    handle="range",
                ),
                *((picker,) if rendered else ()),
                Node(role="AXButton", title="Join with an ID", handle="join-id"),
            ),
        )
        buttons = tuple(
            Node(role="AXButton", title=event.title(), handle=f"event:{index}")
            for index, event in enumerate(self.events)
            if first <= event.day <= last
        )
        grid = Node(
            role="AXGroup",
            subrole="AXLandmarkMain",
            title="calendar view, current time: Sun 14:36",
            children=buttons,
        )
        layers = [toolbar, grid]
        if self.view_menu:
            layers.append(menu("view", *VIEWS))
        if self.peek_for is not None:
            layers.append(self.peek(self.events[self.peek_for]))
        if self.more_menu:
            layers.append(menu("more", *self.more_items))
        if self.series_menu:
            layers.append(menu("series", "This event", "All events in the series"))
        return layers

    def peek(self, event: DeskEvent) -> Node:
        kids = [
            group("Title", event.subject, Node(role="AXButton", title="View event", handle="view")),
            Node(role="AXButton", title="Join", handle="join"),
            Node(role="AXButton", title="Chat with participants", handle="chat"),
            group(
                "Time", "Mon 11/16/2026 10:00 - 10:30", *(("Series",) if event.recurring else ())
            ),
            group("Location", event.location),
            Node(
                role="AXPopUpButton",
                title=f"Opens card for {event.organizer}",
                handle="organizer",
            ),
            static(f"{event.organizer} invited you." if event.invited else ""),
            static("Accepted 2, Didn't respond 12"),
            group("Notes", event.notes),
        ]
        if event.invited:
            kids.append(self.respond_form(event))
        return Node(role="AXGroup", classes=("ms-Callout", "ms-Callout-main"), children=tuple(kids))

    def respond_form(self, event: DeskEvent) -> Node:
        accept = RESPONSE_LABELS.get(event.response, "Accept")
        return Node(
            role="AXGroup",
            title="respond form",
            children=(
                Node(
                    role="AXCheckBox",
                    title="Email organizer",
                    value="1" if self.email_checked else "0",
                    handle="email",
                ),
                Node(
                    role="AXTextArea",
                    title=MESSAGE_PLACEHOLDER,
                    value=self.message or MESSAGE_PLACEHOLDER,
                    handle="message",
                ),
                Node(role="AXButton", title=accept, dom_id="Accept", handle="accept"),
                Node(role="AXButton", title="Decline", dom_id="Decline", handle="decline"),
                Node(role="AXButton", title="Follow", handle="follow"),
                Node(
                    role="AXButton",
                    title="More options",
                    dom_id="calendar_ellipses_menu",
                    handle="more",
                ),
            ),
        )

    def press(self, node: Node) -> None:
        handle = str(node.handle)
        self.pressed.append(handle)
        if handle.startswith("rail:"):
            self.switch(handle.removeprefix("rail:"))
        elif handle.startswith("view:"):
            self.view, self.view_menu = handle.removeprefix("view:"), False
        elif handle.startswith("event:"):
            self.peek_for = int(handle.removeprefix("event:"))
        else:
            self.press_control(handle)

    def press_control(self, handle: str) -> None:
        steps = {"prev": -7, "next": 7}
        if handle in steps:
            self.week_first += timedelta(days=steps[handle])
        elif handle == "today":
            self.week_first = week_start(TODAY)
        elif handle == "picker":
            self.view_menu = True
        elif handle == "email":
            self.email_checked = not self.email_checked
        elif handle == "more":
            self.more_menu = True
        elif handle in ("accept", "decline", "item:Tentative"):
            self.commit(handle)

    def commit(self, handle: str) -> None:
        self.more_menu = False
        if self.series_prompt:
            self.series_menu = True
            return
        states = {
            "accept": MyResponse.ACCEPTED,
            "decline": MyResponse.DECLINED,
            "item:Tentative": MyResponse.TENTATIVE,
        }
        if self.apply_response and self.peek_for is not None:
            self.events[self.peek_for].response = states[handle]
            self.message = ""
        if self.close_on_commit:
            self.peek_for = None

    def switch(self, title: str) -> None:
        self.app = dict(RAIL)[title]
        self.peek_for = None
        if self.app is TeamsApp.CALENDAR:
            self.picker_hidden = self.picker_delay

    def scroll_into_view(self, node: Node) -> None:
        self.pressed.append(f"scroll:{node.handle}")

    def focus(self, node: Node) -> None:
        self.focused.append(str(node.handle))

    def set_value(self, node: Node, text: str) -> None:
        if node.handle == "message":
            self.message = text

    def show_menu(self, node: Node) -> Node:
        raise ElementNotFoundError(f"no menu for {node.handle!r}")

    def close_popups(self) -> int:
        closed = sum((self.view_menu, self.more_menu, self.series_menu))
        self.view_menu = self.more_menu = self.series_menu = False
        if not self.peek_ignores_popups and self.peek_for is not None:
            self.peek_for = None
            closed += 1
        return closed

    def press_key(self, shortcut: Shortcut) -> None:
        self.keys.append(shortcut)
        if shortcut.key is Key.ESCAPE:
            self.peek_for = None

    def element(self, node: Node) -> object:
        if node.handle is None:
            raise ElementNotFoundError("no handle")
        return node.handle


class DeskText:
    def __init__(self, desk: CalendarDesk) -> None:
        self.desk = desk
        self.writes: list[str] = []
        self.cleared: list[str] = []
        self.typo = ""

    def read(self, element: object, placeholder: str = "") -> str:
        value = self.desk.message if element == "message" else ""
        return "" if placeholder and value == placeholder else value

    def write(
        self, element: object, text: str, at_start: bool = False, placeholder: str = ""
    ) -> None:
        self.writes.append(text)
        if element == "message":
            self.desk.message = self.desk.message + (self.typo or text)

    def clear_exact(self, element: object, inserted: str, placeholder: str = "") -> bool:
        if self.read(element, placeholder) != inserted:
            return False
        self.cleared.append(inserted)
        self.desk.message = ""
        return True


def menu(name: str, *titles: str) -> Node:
    items = tuple(
        Node(role="AXMenuItem", title=title, handle=item_handle(name, title)) for title in titles
    )
    return Node(role="AXMenu", children=items, handle=f"menu:{name}")


def item_handle(name: str, title: str) -> str:
    return f"view:{title}" if name == "view" else f"item:{title}"


def group(title: str, *lines: str | Node) -> Node:
    kids = tuple(line if isinstance(line, Node) else static(line) for line in lines)
    return Node(role="AXGroup", title=title, children=kids)


def static(value: str) -> Node:
    return Node(role="AXStaticText", value=value)


@dataclass
class CalendarRig:
    service: CalendarService
    desk: CalendarDesk
    text: DeskText
    guard: RecordingGuard
    workspace: FakeWorkspace
    app: AppRig
    driver: TeamsDriver
    locations: EventLocations


def calendar_rig(
    desk: CalendarDesk,
    recipients: RecipientGuard | None = None,
    guard: RecordingGuard | None = None,
    workspace: FakeWorkspace | None = None,
) -> CalendarRig:
    rig = app_rig(mode=AccessMode.WEB)
    desk_text = DeskText(desk)
    checks = guard or RecordingGuard()
    places = EventLocations()
    desktop = workspace or FakeWorkspace()
    desktop.running.add(rig.app.require_pid())
    waiter = rig.app.waiter
    chats = ChatService(desk, lambda: TODAY, waiter)
    driver = TeamsDriver(
        app=rig.app,
        gateway=desk,
        chats=chats,
        opener=ChatOpener(desk, chats, waiter),
        text=TextWriter(rig.app, RecordingEvents(), FakePasteboard(), waiter),
        workspace=desktop,
        waiter=waiter,
        zone=UTC,
        focus=NO_IDLE_WAIT,
    )
    deps = CalendarDeps(
        gateway=AxCalendar(desk, desk_text, waiter),
        scope=driver.scope,
        waiter=waiter,
        today=lambda: TODAY,
        zone=UTC,
        recipients=recipients or RecipientGuard(None, "48:notes"),
        guard=checks,
    )
    service = CalendarService(deps, places)
    return CalendarRig(service, desk, desk_text, checks, desktop, rig, driver, places)
