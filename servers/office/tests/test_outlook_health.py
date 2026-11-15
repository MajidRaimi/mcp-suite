from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, cast

import pytest

from office_mcp.features.outlook.ax import AxOutlook
from office_mcp.features.outlook.health import parse_today_label, read_health
from office_mcp.features.outlook.service import OutlookService
from office_mcp.shared.ax import AxApp
from office_mcp.shared.errors import OfficeError, OutlookUnavailableError
from office_mcp.shared.node import Node

SYSTEM_NOW = datetime(2026, 11, 15, 13, 44)
SYSTEM_TODAY = SYSTEM_NOW.date()
HEALTHY = ["New Mail; Press space for other actions", "Selected, Today, Sunday, 15 November"]
STALE_ONLY = ["New Mail; Press space for other actions", "Today, Wednesday, 4 November"]
OFFLINE_AND_STALE = [
    "Go Online",
    "New Mail; Press space for other actions",
    "Selected, Today, Wednesday, 4 November",
]


class HealthGateway:
    def __init__(self, labels: list[str], fail: bool = False) -> None:
        self.labels = labels
        self.fail = fail
        self.listed = False
        self.launches: list[bool] = []

    def window_title(self) -> str:
        return "Inbox • jane@example.com"

    def shallow_button_labels(self, launch: bool = True) -> list[str]:
        self.launches.append(launch)
        if self.fail:
            raise OfficeError("The main Outlook mail window is not open")
        return self.labels

    def folder_rows(self) -> list[Node]:
        return [Node(role="AXRow", description="Inbox; 3 unread messages")]

    def message_rows(self) -> list[Node]:
        self.listed = True
        return []

    def select(self, node: Node) -> None:
        raise AssertionError("must not select while Outlook is not working")

    def press(self, node: Node) -> None:
        raise AssertionError("must not press while Outlook is not working")

    def reading_pane(self) -> Node | None:
        return None

    def toolbar_button(self, name: str) -> Node | None:
        return None

    def press_button(self, node: Node) -> None:
        raise AssertionError("must not press while Outlook is not working")


def service(gateway: HealthGateway) -> OutlookService:
    return OutlookService(gateway, lambda: SYSTEM_TODAY, now=lambda: SYSTEM_NOW)


@pytest.mark.parametrize(
    ("label", "expected"),
    [
        ("Selected, Today, Wednesday, 4 November", date(2026, 11, 4)),
        ("Today, Wednesday 4 November", date(2026, 11, 4)),
        ("Today, Wednesday, November 4", date(2026, 11, 4)),
        ("Today, 4 Nov", date(2026, 11, 4)),
        ("Today, Wednesday, 4 Nov.", date(2026, 11, 4)),
        ("Selected, Today, Sunday, 15 November, 2 events", date(2026, 11, 15)),
        ("Today, Thursday, 31 December", date(2026, 12, 31)),
        ("Selected, Today, Friday, 2 January 2027", date(2027, 1, 2)),
        ("Today, Monday, 30 February", None),
        ("Lunch, Today, 12 PM", None),
        ("Tomorrow, Thursday, 5 November", None),
        ("Today", None),
        ("Go Online", None),
    ],
)
def test_reads_the_date_outlook_believes_is_today(label: str, expected: date | None) -> None:
    assert parse_today_label(label, SYSTEM_TODAY) == expected


def test_today_label_without_a_year_picks_the_nearest_year() -> None:
    assert parse_today_label("Today, 31 December", date(2027, 1, 1)) == date(2026, 12, 31)
    assert parse_today_label("Today, 1 January", date(2026, 12, 31)) == date(2027, 1, 1)


def test_healthy_outlook_is_working() -> None:
    health = read_health(HEALTHY, SYSTEM_NOW)

    assert health.working
    assert health.to_record() == {
        "outlook_working": True,
        "outlook_online": True,
        "outlook_today": "2026-11-15",
        "outlook_problems": [],
    }


def test_offline_and_stuck_outlook_reports_both_problems() -> None:
    health = read_health(OFFLINE_AND_STALE, SYSTEM_NOW)
    record = health.to_record()

    assert not health.working
    assert record["outlook_online"] is False
    assert record["outlook_today"] == "2026-11-04"
    offline, stale = record["outlook_problems"]
    assert "Outlook is offline" in offline and "Go Online" in offline
    assert "Wednesday 4 November 2026" in stale and "Sunday 15 November 2026" in stale
    assert health.message().startswith("Outlook is not working right now.")


def test_outlook_stuck_on_yesterday_is_not_working() -> None:
    health = read_health(["Today, Saturday, 14 November"], SYSTEM_NOW)

    assert not health.working
    assert "stuck on Saturday 14 November 2026" in health.problems[0]


def test_yesterday_is_tolerated_only_just_after_midnight() -> None:
    labels = ["Today, Saturday, 14 November"]

    assert read_health(labels, datetime(2026, 11, 15, 0, 5)).working
    assert not read_health(labels, datetime(2026, 11, 15, 0, 15)).working
    assert not read_health(["Today, Monday, 16 November"], SYSTEM_NOW).working


def test_missing_calendar_pane_only_checks_the_offline_banner() -> None:
    assert read_health(["New Mail"], SYSTEM_NOW).working
    assert read_health(["Go Online"], SYSTEM_NOW).shown_today is None
    assert not read_health(["Go Online"], SYSTEM_NOW).working


@pytest.mark.parametrize("labels", [OFFLINE_AND_STALE, STALE_ONLY, ["Go Online"]])
@pytest.mark.parametrize(
    "call",
    [
        lambda s: s.folders(),
        lambda s: s.conversations("Inbox", False, 3, 20),
        lambda s: s.read("budget", None),
    ],
)
def test_every_outlook_tool_refuses_to_answer_from_stale_mail(
    labels: list[str],
    call: Callable[[OutlookService], object],
) -> None:
    gateway = HealthGateway(labels)

    with pytest.raises(OutlookUnavailableError, match="Outlook is not working right now"):
        call(service(gateway))
    assert not gateway.listed
    assert gateway.launches == [True]


def test_working_outlook_lists_as_before() -> None:
    gateway = HealthGateway(HEALTHY)

    assert service(gateway).conversations("Inbox", False, None, 20) == []
    assert gateway.listed
    assert [folder.name for folder in service(gateway).folders()] == ["Inbox"]


def test_status_never_launches_outlook_and_explains_failures() -> None:
    gateway = HealthGateway([], fail=True)

    record = service(gateway).health_record()

    assert gateway.launches == [False]
    assert record["outlook_working"] is False
    assert record["outlook_online"] is None
    assert record["outlook_problems"] == ["The main Outlook mail window is not open"]


@dataclass
class Element:
    role: str
    description: str = ""
    children: list[Element] = field(default_factory=list)


@dataclass
class FakeApp:
    main: Element | None
    launched: list[str] = field(default_factory=list)

    def require_trust(self) -> None:
        return None

    def window(self, matches: Callable[[str], bool]) -> Element:
        self.launched.append("window")
        if self.main is None:
            raise OfficeError("Microsoft Outlook has no matching window")
        return self.main

    def find_window(self, matches: Callable[[str], bool]) -> Element | None:
        return self.main

    def attribute(self, element: Any, name: str) -> Any:
        node: Element = element
        if name == "AXRole":
            return node.role
        if name == "AXDescription":
            return node.description
        if name == "AXChildren":
            return node.children
        return None


def button(label: str) -> Element:
    return Element(role="AXButton", description=label)


def nest(depth: int, leaf: Element) -> Element:
    node = leaf
    for _ in range(depth):
        node = Element(role="AXGroup", children=[node])
    return node


def outlook_window() -> Element:
    message_list = Element(role="AXTable", children=[button("Go Online inside the list")])
    folders = Element(role="AXOutline", children=[button("Today, Monday, 1 January")])
    return Element(
        role="AXWindow",
        children=[
            button("Go Online"),
            nest(2, button("Selected, Today, Wednesday, 4 November")),
            nest(4, button("Too deep")),
            message_list,
            folders,
        ],
    )


def test_scan_reads_shallow_buttons_and_skips_lists_and_deep_nodes() -> None:
    app = FakeApp(outlook_window())

    labels = AxOutlook(cast(AxApp, app)).shallow_button_labels()

    assert sorted(labels) == ["Go Online", "Selected, Today, Wednesday, 4 November"]


def test_scan_for_status_does_not_open_a_missing_window() -> None:
    app = FakeApp(None)

    with pytest.raises(OfficeError, match="main Outlook mail window is not open"):
        AxOutlook(cast(AxApp, app)).shallow_button_labels(launch=False)
    assert app.launched == []
