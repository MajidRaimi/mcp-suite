from __future__ import annotations

import threading
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, tzinfo
from typing import Any

from office_mcp.features.outlook.service import OutlookService
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.messages.service import MessageService
from office_mcp.shared.errors import PlatformError

UNSUPPORTED_PLATFORM = (
    "office-mcp reads Outlook and Teams through the macOS accessibility API, so its tools only "
    "work on macOS with the pyobjc frameworks installed."
)


@dataclass
class Services:
    outlook: OutlookService
    chats: ChatService
    messages: MessageService
    status: Callable[[], dict[str, Any]]
    today: Callable[[], date]
    lock: threading.Lock = field(default_factory=threading.Lock)


def build_services(zone: tzinfo) -> Services:
    from office_mcp.features.outlook.ax import AxOutlook
    from office_mcp.features.teams.ax import AxTeams

    try:
        outlook_ax = AxOutlook()
        teams_ax = AxTeams()
    except ImportError as error:
        raise PlatformError(UNSUPPORTED_PLATFORM) from error
    chats = ChatService(teams_ax, date.today)
    outlook = OutlookService(outlook_ax, date.today)

    def status() -> dict[str, Any]:
        record: dict[str, Any] = {
            "accessibility": outlook_ax.trusted(),
            "outlook_running": outlook_ax.running(),
            "teams_running": teams_ax.running(),
        }
        if record["accessibility"] and record["outlook_running"]:
            record.update(outlook.health_record())
        return record

    return Services(
        outlook=outlook,
        chats=chats,
        messages=MessageService(teams_ax, chats, zone),
        status=status,
        today=date.today,
    )
