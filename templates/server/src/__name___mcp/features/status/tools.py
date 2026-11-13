from __future__ import annotations

from dataclasses import asdict

from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations

from __name___mcp.features.status.service import StatusService
from __name___mcp.shared.owner import Owner

Result = dict[str, object]

STATUS_ANNOTATIONS = ToolAnnotations(
    read_only_hint=True,
    destructive_hint=False,
    idempotent_hint=True,
    open_world_hint=False,
)


def status_description(owner: Owner) -> str:
    return (
        f"Show that the server is running, whose account it serves ({owner.label}) and the "
        "current time in the configured timezone. Call it when a tool fails to check that "
        "the server itself is reachable."
    )


def register_status_tools(server: MCPServer, owner: Owner, service: StatusService) -> None:
    @server.tool(
        name="__name___status",
        description=status_description(owner),
        annotations=STATUS_ANNOTATIONS,
    )
    def status() -> Result:
        return asdict(service.report())
