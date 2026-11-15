from __future__ import annotations

from collections.abc import Callable
from typing import Any, Literal

from office_mcp.features.calendar import descriptions
from office_mcp.features.calendar.rsvp import rsvp_request
from office_mcp.features.calendar.service import CalendarService
from office_mcp.shared.drafts import Channel
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    ToolSpec,
)
from office_mcp.shared.runner import SEND_BUDGET, Deadline, Lazy, Runner

Record = dict[str, Any]
Response = Literal["accept", "tentative", "decline"]


class CalendarTools:
    def __init__(
        self,
        runner: Runner,
        calendar: Lazy[CalendarService],
        owner: Owner,
        rules_text: Callable[[Channel], str] | None = None,
    ) -> None:
        self._runner = runner
        self._calendar = calendar
        self._owner = owner
        self._rules_text = rules_text

    def specs(self) -> list[ToolSpec]:
        return [
            ToolSpec(
                name="calendar_events",
                function=self.calendar_events,
                description=descriptions.calendar_events(self._owner),
                group=ToolGroup.READ,
                effect=Effect.READS,
            ),
            ToolSpec(
                name="calendar_event",
                function=self.calendar_event,
                description=descriptions.calendar_event(self._owner),
                group=ToolGroup.READ,
                effect=Effect.READS,
            ),
            ToolSpec(
                name="calendar_respond",
                function=self.calendar_respond,
                description=EXPERIMENTAL_PREFIX
                + descriptions.calendar_respond(self._owner, self._rules()),
                group=ToolGroup.SEND,
                effect=Effect.SENDS,
                stability=Stability.EXPERIMENTAL,
            ),
        ]

    def _rules(self) -> str:
        return "" if self._rules_text is None else self._rules_text(Channel.RSVP)

    def calendar_events(
        self,
        start: str | None = None,
        days: int = 1,
        include_canceled: bool = False,
        limit: int = 100,
    ) -> list[Record]:
        def work(deadline: Deadline) -> list[Record]:
            events = self._calendar.get().events(start, days, include_canceled, limit, deadline)
            return [event.to_record() for event in events]

        return self._runner.read(work)

    def calendar_event(self, event: str) -> Record:
        return self._runner.read(
            lambda deadline: self._calendar.get().event(event, deadline).to_record()
        )

    def calendar_respond(
        self,
        event: str,
        response: Response,
        notify: bool = True,
        comment: str | None = None,
        dry_run: bool = False,
    ) -> Record:
        def work(deadline: Deadline) -> Record:
            request = rsvp_request(event, response, notify, comment, dry_run)
            return self._calendar.get().respond(request, deadline).to_record()

        return self._runner.write(work, SEND_BUDGET)
