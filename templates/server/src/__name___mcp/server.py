from __future__ import annotations

import argparse
import os
import sys
from collections.abc import Sequence
from dataclasses import replace

import uvicorn
from mcp.server.mcpserver import MCPServer
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse

from __name___mcp.descriptions import server_instructions
from __name___mcp.features.status.service import StatusService
from __name___mcp.features.status.tools import register_status_tools
from __name___mcp.shared.errors import ConfigurationError
from __name___mcp.shared.http import BearerTokenMiddleware, transport_security
from __name___mcp.shared.owner import Owner
from __name___mcp.shared.settings import Settings, Transport, parse_port

SERVER_NAME = "__name__"
MCP_PATH = "/mcp"


def build_server(settings: Settings) -> MCPServer:
    owner = Owner.from_name(settings.owner_name)
    server = MCPServer(name=SERVER_NAME, instructions=server_instructions(owner))
    register_status_tools(server, owner, StatusService(owner.label, settings.timezone))

    async def health(_: Request) -> JSONResponse:
        return JSONResponse({"status": "ok"})

    server.custom_route("/health", methods=["GET"], include_in_schema=False)(health)
    return server


def build_http_app(server: MCPServer, settings: Settings) -> Starlette:
    app = server.streamable_http_app(
        streamable_http_path=MCP_PATH,
        stateless_http=True,
        json_response=True,
        transport_security=transport_security(settings.allowed_hosts),
        host=settings.host,
    )
    if settings.token:
        app.add_middleware(BearerTokenMiddleware, token=settings.token)
    return app


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog=f"{SERVER_NAME}-mcp")
    parser.add_argument("--transport", choices=[transport.value for transport in Transport])
    parser.add_argument("--host")
    parser.add_argument("--port")
    return parser.parse_args(argv)


def apply_overrides(settings: Settings, args: argparse.Namespace) -> Settings:
    return replace(
        settings,
        transport=Transport(args.transport) if args.transport else settings.transport,
        host=args.host or settings.host,
        port=parse_port(args.port) if args.port else settings.port,
    )


def serve(settings: Settings) -> None:
    server = build_server(settings)
    if settings.transport is Transport.STDIO:
        server.run("stdio")
        return
    uvicorn.run(build_http_app(server, settings), host=settings.host, port=settings.port)


def main(argv: Sequence[str] | None = None) -> None:
    try:
        settings = apply_overrides(Settings.from_env(os.environ), parse_args(argv))
    except ConfigurationError as error:
        sys.exit(f"{SERVER_NAME}-mcp: {error}")
    serve(settings)
