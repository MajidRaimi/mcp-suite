from __future__ import annotations

import argparse
import sys
from collections.abc import Callable, Sequence
from dataclasses import replace

import uvicorn
from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse

from office_mcp import descriptions as text
from office_mcp.container import build_services
from office_mcp.shared.errors import SettingsError
from office_mcp.shared.http import BearerTokenMiddleware, transport_security
from office_mcp.shared.settings import TRANSPORTS, Settings, parse_port, parse_transport
from office_mcp.tools import OfficeTools

STREAMABLE_HTTP_PATH = "/mcp"

READ_ONLY = ToolAnnotations(
    read_only_hint=True,
    destructive_hint=False,
    idempotent_hint=True,
    open_world_hint=False,
)
MARKS_READ = ToolAnnotations(
    read_only_hint=False,
    destructive_hint=False,
    idempotent_hint=False,
    open_world_hint=False,
)


def build_tools(settings: Settings) -> OfficeTools:
    return OfficeTools(lambda: build_services(settings.zone.info))


def create_server(settings: Settings, tools: OfficeTools) -> MCPServer:
    owner = text.Owner(settings.owner_name)
    server = MCPServer("office", instructions=text.instructions(owner, settings.zone.name))
    register: list[tuple[Callable[..., object], str, ToolAnnotations]] = [
        (tools.office_status, text.OFFICE_STATUS, READ_ONLY),
        (tools.outlook_folders, text.OUTLOOK_FOLDERS, READ_ONLY),
        (tools.outlook_conversations, text.OUTLOOK_CONVERSATIONS, READ_ONLY),
        (tools.outlook_read, text.OUTLOOK_READ, MARKS_READ),
        (tools.teams_chats, text.TEAMS_CHATS, READ_ONLY),
        (tools.teams_read, text.TEAMS_READ, MARKS_READ),
        (tools.teams_needs_reply, text.teams_needs_reply(owner), READ_ONLY),
    ]
    for function, description, hints in register:
        server.tool(name=function.__name__, description=description, annotations=hints)(function)
    server.custom_route("/health", methods=["GET"], include_in_schema=False)(health)
    return server


async def health(_request: Request) -> JSONResponse:
    return JSONResponse({"status": "ok"})


def http_app(server: MCPServer, settings: Settings) -> Starlette:
    app = server.streamable_http_app(
        streamable_http_path=STREAMABLE_HTTP_PATH,
        json_response=True,
        stateless_http=True,
        transport_security=transport_security(settings.allowed_hosts),
        host=settings.host,
    )
    if settings.token:
        app.add_middleware(BearerTokenMiddleware, token=settings.token)
    return app


def serve(server: MCPServer, settings: Settings) -> None:
    if settings.transport == "stdio":
        server.run("stdio")
        return
    uvicorn.run(http_app(server, settings), host=settings.host, port=settings.port)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="office-mcp",
        description="Serve the Outlook and Teams tools over MCP.",
    )
    parser.add_argument(
        "--transport",
        choices=TRANSPORTS,
        help="stdio (default) or streamable-http; overrides OFFICE_MCP_TRANSPORT",
    )
    parser.add_argument(
        "--http",
        dest="transport",
        action="store_const",
        const="streamable-http",
        help="shorthand for --transport streamable-http",
    )
    parser.add_argument("--host", help="HTTP bind address; overrides OFFICE_MCP_HOST")
    parser.add_argument("--port", help="HTTP port; overrides OFFICE_MCP_PORT")
    return parser


def apply_arguments(settings: Settings, args: argparse.Namespace) -> Settings:
    return replace(
        settings,
        transport=parse_transport(args.transport, "--transport")
        if args.transport
        else settings.transport,
        host=args.host or settings.host,
        port=parse_port(args.port, "--port", settings.port),
    )


def main(argv: Sequence[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    try:
        settings = apply_arguments(Settings.from_env(), args)
    except SettingsError as error:
        print(f"office-mcp: {error}", file=sys.stderr)
        raise SystemExit(2) from error
    serve(create_server(settings, build_tools(settings)), settings)


if __name__ == "__main__":
    main()
