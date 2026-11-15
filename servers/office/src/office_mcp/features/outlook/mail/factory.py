from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from office_mcp.features.outlook.driver import OutlookDriver
from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.mail.attachments import AttachmentService
from office_mcp.features.outlook.mail.downloads import PathDownloadsWatcher
from office_mcp.features.outlook.mail.search import AxSearchScope, SearchService
from office_mcp.features.outlook.mail.service import MailService
from office_mcp.features.outlook.people.service import AxSuggestions, PeopleService
from office_mcp.features.outlook.service import MenuUnreadMarker, OutlookService
from office_mcp.shared.runtime import Runtime

DOWNLOADS = Path.home() / "Downloads"


def main_window(gateway: OutlookGateway) -> Callable[[], object]:
    def window() -> object:
        return gateway.main_probe().handle

    return window


def owner_address(runtime: Runtime, gateway: OutlookGateway) -> Callable[[], str | None]:
    def address() -> str | None:
        return runtime.settings.owner_email or gateway.account()

    return address


def build_attachments(
    core: OutlookService, outlook: OutlookDriver, directory: Path | None
) -> AttachmentService | None:
    if directory is None:
        return None
    waiter = outlook.waiter
    watcher = PathDownloadsWatcher(DOWNLOADS, waiter.clock, waiter.sleep)
    return AttachmentService(core, outlook.gateway, outlook.popups, watcher, directory, waiter)


def build_mail(runtime: Runtime, outlook: OutlookDriver) -> MailService:
    gateway = outlook.gateway
    waiter = outlook.waiter
    window = main_window(gateway)
    core = OutlookService(
        gateway,
        outlook.rows,
        outlook.scope,
        waiter,
        runtime.today,
        runtime.now,
        MenuUnreadMarker(outlook.menus),
    )
    search = SearchService(
        core, outlook.rows, gateway, AxSearchScope(outlook.app, window, outlook.popups, waiter)
    )
    people = PeopleService(
        AxSuggestions(outlook.app, window, outlook.events, waiter), outlook.scope, waiter
    )
    return MailService(
        core,
        search,
        people,
        build_attachments(core, outlook, runtime.settings.download_dir),
        owner_address(runtime, gateway),
        runtime.today,
    )
