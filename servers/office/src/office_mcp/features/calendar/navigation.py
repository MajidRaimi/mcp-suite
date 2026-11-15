from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date

from office_mcp.features.calendar.gateway import CalendarGateway, CalendarScope
from office_mcp.features.calendar.grid import (
    NEXT_WEEK_PREFIX,
    PREVIOUS_WEEK_PREFIX,
    TODAY_PREFIX,
    WEEK_VIEW,
    calendar_view,
    shown_range,
    shown_view,
    toolbar_button,
    view_option,
    view_picker,
)
from office_mcp.features.calendar.models import WEEK_DAYS, DateRange
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.errors import ElementNotFoundError, TargetError, UiTimeoutError
from office_mcp.shared.node import Node
from office_mcp.shared.runner import Deadline
from office_mcp.shared.waiting import Waiter

CALENDAR_LOAD_WAIT = 12.0
MAX_WEEK_STEPS = 60
LOAD_TIMEOUT = "Teams calendar did not load in 12 s"
RESTORE_VIEW = "restore the Teams calendar view"
RESTORE_WEEK = "show the Teams calendar week from before"
OUT_OF_TIME = "the Teams calendar took too long; ask for fewer days"


@dataclass(frozen=True)
class CalendarPage:
    root: Node
    range: DateRange


class CalendarNavigator:
    def __init__(
        self,
        gateway: CalendarGateway,
        waiter: Waiter,
        scope: CalendarScope,
        deadline: Deadline,
        today: Callable[[], date],
    ) -> None:
        self._gateway = gateway
        self._waiter = waiter
        self._scope = scope
        self._deadline = deadline
        self._today = today
        self._week_restore_pushed = False

    def open(self) -> CalendarPage:
        self._scope.show_app(TeamsApp.CALENDAR)
        root = self._waiter.require(
            self._ready, lambda: UiTimeoutError(LOAD_TIMEOUT), CALENDAR_LOAD_WAIT
        )
        return self._page(self._ensure_week_view(root))

    def show_week(self, page: CalendarPage, day: date) -> CalendarPage:
        origin = page.range
        steps = 0
        while not page.range.contains(day):
            if steps >= MAX_WEEK_STEPS:
                raise TargetError(f"{day.isoformat()} is too far from the week Teams shows")
            self._require_time()
            self._push_week_restore(origin)
            page = self._step(page, forward=day > page.range.last)
            steps += 1
        return page

    def next_week(self, page: CalendarPage) -> CalendarPage:
        self._require_time()
        self._push_week_restore(page.range)
        return self._step(page, forward=True)

    def _ready(self) -> Node | None:
        root = self._gateway.snapshot()
        if calendar_view(root) is None or view_picker(root) is None or shown_range(root) is None:
            return None
        return root

    def _page(self, root: Node) -> CalendarPage:
        shown = shown_range(root)
        if shown is None:
            raise ElementNotFoundError("could not read which week the Teams calendar shows")
        return CalendarPage(root, shown)

    def _ensure_week_view(self, root: Node) -> Node:
        previous = shown_view(root)
        if previous is None:
            raise ElementNotFoundError("the Teams calendar view picker was not found")
        if previous == WEEK_VIEW:
            return root
        self._scope.cleanup.push(RESTORE_VIEW, lambda: self._restore_view(previous))
        return self._choose_view(WEEK_VIEW)

    def _restore_view(self, view: str) -> None:
        self._choose_view(view)

    def _choose_view(self, view: str) -> Node:
        root = self._gateway.snapshot()
        if shown_view(root) == view:
            return root
        picker = view_picker(root)
        if picker is None:
            raise ElementNotFoundError("the Teams calendar view picker was not found")
        self._gateway.press(picker)
        option = self._waiter.require(
            lambda: view_option(self._gateway.snapshot(), view),
            lambda: ElementNotFoundError(f"the Teams calendar has no {view!r} view option"),
        )
        self._gateway.press(option)
        return self._waiter.require(
            lambda: self._showing_view(view),
            lambda: UiTimeoutError(f"the Teams calendar did not switch to the {view} view"),
        )

    def _showing_view(self, view: str) -> Node | None:
        root = self._gateway.snapshot()
        if shown_view(root) != view or shown_range(root) is None:
            return None
        return root

    def _step(self, page: CalendarPage, forward: bool) -> CalendarPage:
        prefix = NEXT_WEEK_PREFIX if forward else PREVIOUS_WEEK_PREFIX
        button = toolbar_button(page.root, prefix)
        if button is None:
            raise ElementNotFoundError("the Teams calendar week buttons were not found")
        expected = page.range.shifted(WEEK_DAYS if forward else -WEEK_DAYS)
        self._gateway.press(button)
        root = self._waiter.require(
            lambda: self._showing_range(expected),
            lambda: UiTimeoutError(
                f"the Teams calendar did not show the week {expected.describe()}"
            ),
        )
        return CalendarPage(root, expected)

    def _showing_range(self, expected: DateRange) -> Node | None:
        root = self._gateway.snapshot()
        return root if shown_range(root) == expected else None

    def _require_time(self) -> None:
        if self._deadline.expired():
            raise UiTimeoutError(OUT_OF_TIME)

    def _push_week_restore(self, origin: DateRange) -> None:
        if self._week_restore_pushed:
            return
        self._week_restore_pushed = True
        self._scope.cleanup.push(RESTORE_WEEK, lambda: self._restore_week(origin))

    def _restore_week(self, origin: DateRange) -> None:
        root = self._gateway.snapshot()
        current = shown_range(root)
        if current == origin:
            return
        if origin.contains(self._today()):
            self._go_to_today(root, origin)
            return
        page = self._page(root)
        steps = 0
        while page.range != origin and steps < MAX_WEEK_STEPS:
            page = self._step(page, forward=origin.first > page.range.first)
            steps += 1

    def _go_to_today(self, root: Node, origin: DateRange) -> None:
        button = toolbar_button(root, TODAY_PREFIX)
        if button is None:
            raise ElementNotFoundError("the Teams calendar has no Go to today button")
        self._gateway.press(button)
        self._waiter.require(
            lambda: self._showing_range(origin),
            lambda: UiTimeoutError("the Teams calendar did not go back to this week"),
        )
