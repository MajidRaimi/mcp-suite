from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.mail.downloads import DownloadsWatcher, move_into, size_matches
from office_mcp.features.outlook.reading import AttachmentEntry, attachment_entries
from office_mcp.features.outlook.service import OutlookService
from office_mcp.shared.errors import AttachmentError
from office_mcp.shared.menus import PopupMenus
from office_mcp.shared.node import Node
from office_mcp.shared.waiting import Waiter

DOWNLOAD_ITEM = ("Download",)
DOWNLOAD_TIMEOUT = 30.0
ATTACHMENT_WAIT = 4.0
PANE_BODY_LIMIT = 2000


@dataclass(frozen=True)
class SavedAttachment:
    path: Path
    name: str
    size_bytes: int
    notes: tuple[str, ...] = ()

    def to_record(self) -> dict[str, Any]:
        return {
            "path": str(self.path),
            "name": self.name,
            "size_bytes": self.size_bytes,
            "notes": list(self.notes),
        }


def named_entries(pane: Node | None, name: str) -> list[AttachmentEntry] | None:
    if pane is None:
        return None
    found = [entry for entry in attachment_entries(pane) if entry.attachment.name == name]
    return found or None


def missing_error(pane: Node, name: str) -> AttachmentError:
    names = ", ".join(entry.attachment.name for entry in attachment_entries(pane)) or "none"
    return AttachmentError(f"no attachment named {name!r}; attachments: {names}")


class AttachmentService:
    def __init__(
        self,
        core: OutlookService,
        gateway: OutlookGateway,
        popups: PopupMenus,
        watcher: DownloadsWatcher,
        directory: Path,
        waiter: Waiter,
        timeout: float = DOWNLOAD_TIMEOUT,
    ) -> None:
        self._core = core
        self._gateway = gateway
        self._popups = popups
        self._watcher = watcher
        self._directory = directory
        self._waiter = waiter
        self._timeout = timeout

    def save(self, ref: str, name: str, folder: str | None, keep_unread: bool) -> SavedAttachment:
        with self._core.scope(foreground=True) as ui:
            conversation = self._core.open(
                ui, ref, folder, keep_unread, PANE_BODY_LIMIT, 0, keys_only=True
            )
            entry = self._entry(conversation.opened.pane, name)
            downloaded = self._download(entry)
            saved = move_into(downloaded, self._directory, entry.attachment.name)
            notes = conversation.close()
        return SavedAttachment(saved, entry.attachment.name, saved.stat().st_size, notes)

    def _entry(self, pane: Node, name: str) -> AttachmentEntry:
        found = self._waiter.until(
            lambda: named_entries(self._gateway.reading_pane(), name), ATTACHMENT_WAIT
        )
        if found is None:
            raise missing_error(pane, name)
        if len(found) > 1:
            raise AttachmentError(f"several attachments are named {name!r}; none was saved")
        return found[0]

    def _download(self, entry: AttachmentEntry) -> Path:
        button = entry.menu_button
        if button is None or button.handle is None:
            raise AttachmentError(f"the attachment {entry.attachment.name!r} has no menu")
        before = self._watcher.snapshot()
        menu = self._popups.open(button.handle, "AXShowMenu")
        try:
            self._popups.choose(menu, DOWNLOAD_ITEM)
        finally:
            self._popups.close()
        downloaded = self._watcher.wait_for(entry.attachment.name, before, self._timeout)
        actual = downloaded.stat().st_size
        if not size_matches(entry.attachment.size, actual):
            raise AttachmentError(
                f"the downloaded file {downloaded.name!r} has {actual} bytes, but Outlook "
                f"shows {entry.attachment.size}; it was left in {downloaded.parent}"
            )
        return downloaded
