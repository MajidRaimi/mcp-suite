from __future__ import annotations

from office_mcp.features.outlook.actions.ax import AxActions
from office_mcp.features.outlook.actions.service import ActionsService
from office_mcp.features.outlook.driver import OutlookDriver
from office_mcp.shared.runtime import Runtime


def build_actions(runtime: Runtime, outlook: OutlookDriver) -> ActionsService:
    gateway = AxActions(
        outlook.app,
        outlook.windows,
        outlook.menus,
        outlook.popups,
        outlook.events,
        outlook.waiter,
    )
    return ActionsService(
        outlook=outlook.gateway,
        gateway=gateway,
        rows=outlook.rows,
        scopes=outlook.scope,
        waiter=outlook.waiter,
        today=runtime.today,
    )
