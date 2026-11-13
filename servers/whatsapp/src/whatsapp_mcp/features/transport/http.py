from __future__ import annotations

from mcp.server.mcpserver import MCPServer
from starlette.applications import Starlette

from whatsapp_mcp.features.transport.auth import BearerTokenMiddleware
from whatsapp_mcp.features.transport.security import transport_security
from whatsapp_mcp.shared.settings import Settings

MCP_PATH = "/mcp"


def build_http_app(server: MCPServer, settings: Settings) -> Starlette:
    app = server.streamable_http_app(
        streamable_http_path=MCP_PATH,
        json_response=True,
        stateless_http=True,
        transport_security=transport_security(settings.allowed_hosts),
        host=settings.host,
    )
    if settings.token:
        app.add_middleware(BearerTokenMiddleware, token=settings.token)
    return app
