from __future__ import annotations

from collections.abc import Mapping
from enum import StrEnum

from office_mcp.features.outlook.actions.gateway import ActionsGateway
from office_mcp.features.outlook.actions.neighbors import UnreadKeeper
from office_mcp.features.outlook.actions.watch import RowWatch, Target, is_flagged, is_unflagged
from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.shared.results import ItemResult

FLAG_MENU = "Flag"
CLEAR_ITEM = "Clear Flag"
VERDICT_NOTES: Mapping[bool | None, str] = {
    False: "Outlook does not show the flag change",
    None: "the flag could not be read back",
}


class FlagState(StrEnum):
    TODAY = "today"
    TOMORROW = "tomorrow"
    THIS_WEEK = "this_week"
    NEXT_WEEK = "next_week"
    NO_DATE = "no_date"
    CLEAR = "clear"


DUE_ITEMS: Mapping[FlagState, str] = {
    FlagState.TODAY: "Today",
    FlagState.TOMORROW: "Tomorrow",
    FlagState.THIS_WEEK: "This Week",
    FlagState.NEXT_WEEK: "Next Week",
    FlagState.NO_DATE: "No Date",
}


def menu_path(state: FlagState) -> tuple[str, str]:
    return (FLAG_MENU, DUE_ITEMS.get(state, CLEAR_ITEM))


def wants_flag(state: FlagState) -> bool:
    return state is not FlagState.CLEAR


def checked_verdict(state: FlagState, checks: Mapping[str, bool | None]) -> bool | None:
    if wants_flag(state):
        return checks.get(DUE_ITEMS[state])
    values = [checks.get(item) for item in DUE_ITEMS.values()]
    if any(value is True for value in values):
        return False
    return True if values and all(value is False for value in values) else None


class Flagger:
    def __init__(
        self,
        outlook: OutlookGateway,
        gateway: ActionsGateway,
        watch: RowWatch,
        keeper: UnreadKeeper,
    ) -> None:
        self._outlook = outlook
        self._gateway = gateway
        self._watch = watch
        self._keeper = keeper

    def apply(self, target: Target, state: FlagState) -> ItemResult:
        summary = target.row.summary
        if state is FlagState.CLEAR and is_unflagged(summary):
            return self._result(target, verified=True, detail="already not flagged")
        self._keeper.select(target.row)
        self._choose(menu_path(state))
        verdict = self._verify(target.key, state, summary.flagged)
        notes = [VERDICT_NOTES[verdict]] if verdict is not True else []
        if summary.unread and not self._keeper.mark_unread(target.row):
            notes.append("Outlook marked it read and it could not be marked unread again")
        self._outlook.deselect(target.row.node)
        detail = "; ".join([f"flag set to {state.value}", *notes])
        return self._result(
            target, verified=verdict is True, detail=detail, ok=verdict is not False
        )

    def _choose(self, path: tuple[str, ...]) -> None:
        menu = self._gateway.open_row_menu()
        try:
            self._gateway.choose(menu, path)
        finally:
            self._gateway.close_menu(menu)

    def _verify(self, key: str, state: FlagState, was_flagged: bool) -> bool | None:
        if wants_flag(state) and was_flagged:
            return checked_verdict(state, self._menu_checks())
        check = is_flagged if wants_flag(state) else is_unflagged
        if self._watch.wait_for(key, check):
            return True
        return checked_verdict(state, self._menu_checks())

    def _menu_checks(self) -> dict[str, bool | None]:
        menu = self._gateway.open_row_menu()
        try:
            return {
                item: self._gateway.checked(menu, (FLAG_MENU, item)) for item in DUE_ITEMS.values()
            }
        finally:
            self._gateway.close_menu(menu)

    @staticmethod
    def _result(target: Target, verified: bool, detail: str, ok: bool = True) -> ItemResult:
        return ItemResult(
            ref=target.key,
            ok=ok,
            label=target.label,
            location=target.folder,
            verified=verified,
            detail=detail,
        )
