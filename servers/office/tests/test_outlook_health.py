from __future__ import annotations

from collections.abc import Callable
from datetime import date, datetime

import pytest
from fakes.ax_backend import FakeElement, element
from fakes.clock import fake_waiter
from fakes.events import RecordingEvents
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.scope import checked_factory
from fakes.trees import app_rig, folder_row, outlook_main_window, window

from office_mcp.features.outlook.ax import AxOutlook
from office_mcp.features.outlook.health import parse_today_label, read_health, require_working
from office_mcp.features.outlook.rows import RefRegistry, RowFinder
from office_mcp.features.outlook.service import OutlookService
from office_mcp.shared.ax import AxApp
from office_mcp.shared.errors import OfficeError, OutlookUnavailableError
from office_mcp.shared.menus import PopupMenus
from office_mcp.shared.windows import WindowLocator

SYSTEM_NOW = datetime(2026, 11, 15, 13, 44)
SYSTEM_TODAY = SYSTEM_NOW.date()
HEALTHY = ["New Mail; Press space for other actions", "Selected, Today, Sunday, 15 November"]
STALE_ONLY = ["New Mail; Press space for other actions", "Today, Wednesday, 4 November"]
OFFLINE_AND_STALE = [
    "Go Online",
    "New Mail; Press space for other actions",
    "Selected, Today, Wednesday, 4 November",
]


def health_gateway(labels: list[str]) -> FakeOutlookGateway:
    return FakeOutlookGateway(
        folders=[folder_row("Inbox; 3 unread messages")], labels=tuple(labels)
    )


def service(gateway: FakeOutlookGateway) -> OutlookService:
    waiter = fake_waiter()
    rows = RowFinder(gateway, waiter, lambda: SYSTEM_TODAY, RefRegistry())

    def check() -> None:
        require_working(gateway.shallow_button_labels(), SYSTEM_NOW)

    scopes, _ = checked_factory(check)
    return OutlookService(gateway, rows, scopes, waiter, lambda: SYSTEM_TODAY, lambda: SYSTEM_NOW)


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
    gateway = health_gateway(labels)

    with pytest.raises(OutlookUnavailableError, match="Outlook is not working right now"):
        call(service(gateway))
    assert gateway.calls == []
    assert gateway.launches == [True]


def test_working_outlook_lists_as_before() -> None:
    gateway = health_gateway(HEALTHY)

    assert service(gateway).conversations("Inbox", False, None, 20) == []
    assert "message_rows" in gateway.calls
    assert [folder.name for folder in service(gateway).folders()] == ["Inbox"]


def test_status_never_launches_outlook_and_explains_failures() -> None:
    gateway = health_gateway([])
    gateway.fail_labels = "The main Outlook mail window is not open"

    record = service(gateway).health_record()

    assert gateway.launches == [False]
    assert record["outlook_working"] is False
    assert record["outlook_online"] is None
    assert record["outlook_problems"] == ["The main Outlook mail window is not open"]


def ax_outlook(app: AxApp) -> AxOutlook:
    events = RecordingEvents()
    return AxOutlook(app, WindowLocator(app), PopupMenus(app, events, fake_waiter()), events)


def button(label: str) -> FakeElement:
    return element("AXButton", AXDescription=label)


def nest(depth: int, leaf: FakeElement) -> FakeElement:
    node = leaf
    for _ in range(depth):
        node = element("AXGroup", node)
    return node


def test_scan_reads_shallow_buttons_and_skips_lists_and_deep_nodes() -> None:
    main = outlook_main_window(
        button("Go Online"),
        nest(2, button("Selected, Today, Wednesday, 4 November")),
        nest(4, button("Too deep")),
        element("AXTable", button("Go Online inside the list")),
        element("AXOutline", button("Today, Monday, 1 January")),
    )
    rig = app_rig(main)

    labels = ax_outlook(rig.app).shallow_button_labels()

    assert sorted(labels) == ["Go Online", "Selected, Today, Wednesday, 4 November"]


def test_scan_for_status_does_not_open_a_missing_window() -> None:
    rig = app_rig(window("Untitled • jane@example.com", AXSubrole="AXStandardWindow"))

    with pytest.raises(OfficeError, match="main Outlook mail window is not open"):
        ax_outlook(rig.app).shallow_button_labels(launch=False)
    assert rig.processes.launched == []
