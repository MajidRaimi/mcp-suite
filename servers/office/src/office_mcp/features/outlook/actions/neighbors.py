from __future__ import annotations

from collections.abc import Sequence

from office_mcp.features.outlook.actions.gateway import ActionsGateway
from office_mcp.features.outlook.actions.watch import RowWatch, is_read, is_unread
from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.models import Row
from office_mcp.shared.errors import ActionNotVerifiedError

MARK_UNREAD = ("Message", "Mark as Unread")
MARK_READ = ("Message", "Mark as Read")
FIRST_SELECT_WAIT = 1.5


def adjacent_unread(rows: Sequence[Row], key: str) -> tuple[str, ...]:
    keys = [row.summary.key for row in rows]
    if key not in keys:
        return ()
    index = keys.index(key)
    around = [rows[index - 1]] if index > 0 else []
    around += [rows[index + 1]] if index + 1 < len(rows) else []
    return tuple(row.summary.key for row in around if row.summary.unread)


def turned_read(watched: Sequence[str], rows: Sequence[Row]) -> list[Row]:
    return [row for row in rows if row.summary.key in watched and not row.summary.unread]


class UnreadKeeper:
    def __init__(self, outlook: OutlookGateway, gateway: ActionsGateway, watch: RowWatch) -> None:
        self._outlook = outlook
        self._gateway = gateway
        self._watch = watch

    def neighbors(self, key: str) -> tuple[str, ...]:
        return adjacent_unread(self._watch.current(), key)

    def restore_neighbors(self, watched: Sequence[str]) -> tuple[str, ...]:
        if not watched:
            return ()
        notes: list[str] = []
        for row in turned_read(watched, self._watch.current()):
            restored = self.mark_unread(row)
            self._outlook.deselect(row.node)
            title = row.summary.title[:60]
            notes.append(
                f"marked the neighboring {title!r} unread again"
                if restored
                else f"Outlook marked the neighboring {title!r} read and it could not be undone"
            )
        return tuple(notes)

    def select(self, row: Row) -> None:
        self._clear_others(row)
        self._outlook.select(row.node)
        if not self._watch.wait_selected(row.summary.key, FIRST_SELECT_WAIT):
            shown = self._watch.find(row.summary.key)
            self._outlook.select(shown.node if shown is not None else row.node)
        self.require_selected(row)

    def require_selected(self, row: Row) -> None:
        if not self._watch.wait_selected(row.summary.key) or self._watch.multi_selection_shown():
            raise ActionNotVerifiedError(
                f"Outlook did not select only {row.summary.title[:60]!r}, so nothing was changed"
            )

    def _clear_others(self, row: Row) -> None:
        for other in self._watch.selected_rows():
            if other.summary.key != row.summary.key:
                self._outlook.deselect(other.node)

    def mark_unread(self, row: Row) -> bool:
        self.select(row)
        key = row.summary.key
        for _ in range(2):
            if not self._offers(MARK_UNREAD):
                return self._watch.wait_for(key, is_unread)
            self._gateway.press_menu(MARK_UNREAD)
            if self._watch.wait_for(key, is_unread):
                return True
        return False

    def _offers(self, path: tuple[str, ...]) -> bool:
        return path[-1] in self._gateway.menu_titles(path[:-1])

    def mark_read(self, row: Row) -> bool:
        self.select(row)
        key = row.summary.key
        if self._watch.wait_for(key, is_read):
            return True
        self._gateway.press_menu(MARK_READ)
        return self._watch.wait_for(key, is_read)
