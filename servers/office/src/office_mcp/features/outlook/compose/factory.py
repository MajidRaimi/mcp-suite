from __future__ import annotations

import time
from collections.abc import Callable

from office_mcp.features.outlook.compose.ax import AxCompose
from office_mcp.features.outlook.compose.guard import SendGuard
from office_mcp.features.outlook.compose.service import ComposeDeps, ComposeService
from office_mcp.features.outlook.conversation import ConversationReader
from office_mcp.features.outlook.driver import OutlookDriver
from office_mcp.features.outlook.window import OutlookWindow, account_from_title, classify_window
from office_mcp.shared.ax import AxApp
from office_mcp.shared.drafts import TextGuard
from office_mcp.shared.runtime import Runtime
from office_mcp.shared.windows import WindowLocator


def owner_address(
    configured: str | None, app: AxApp, windows: WindowLocator
) -> Callable[[], str | None]:
    def address() -> str | None:
        if configured:
            return configured
        if not app.running():
            return None
        main = windows.first(classify_window, OutlookWindow.MAIN)
        return account_from_title(main.title) if main is not None else None

    return address


def build_compose(
    runtime: Runtime, outlook: OutlookDriver, reader: ConversationReader, guard: TextGuard
) -> ComposeService:
    gateway = AxCompose(
        outlook.app,
        outlook.windows,
        outlook.menus,
        outlook.popups,
        outlook.text,
        outlook.gateway,
        outlook.waiter,
    )
    return ComposeService(
        ComposeDeps(
            gateway=gateway,
            rows=outlook.rows,
            scope=outlook.scope,
            reader=reader,
            guard=guard,
            recipients=runtime.recipients,
            sends=SendGuard(time.monotonic),
            waiter=outlook.waiter,
            now=runtime.now,
            owner_email=owner_address(runtime.settings.owner_email, outlook.app, outlook.windows),
        )
    )
