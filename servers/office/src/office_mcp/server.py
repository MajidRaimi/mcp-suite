from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from dataclasses import replace

import uvicorn
from mcp.server.mcpserver import MCPServer
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse

from office_mcp import descriptions as text
from office_mcp.container import Container, build_container, exposed_specs
from office_mcp.features.persona.render import instructions_block
from office_mcp.shared.errors import SettingsError
from office_mcp.shared.http import BearerTokenMiddleware, transport_security
from office_mcp.shared.registry import ToolGroup, ToolSpec, annotations_for
from office_mcp.shared.settings import TRANSPORTS, Settings, parse_port, parse_transport

STREAMABLE_HTTP_PATH = "/mcp"


def server_instructions(container: Container) -> str:
    settings = container.settings
    groups = settings.tool_policy.groups
    block = (
        instructions_block(container.persona, container.owner) if ToolGroup.SEND in groups else ""
    )
    return text.instructions(container.owner, settings.zone.name, groups, block)


def register(server: MCPServer, spec: ToolSpec) -> None:
    server.tool(
        name=spec.name,
        description=spec.description,
        annotations=annotations_for(spec.effect),
    )(spec.function)


def create_server(container: Container) -> MCPServer:
    specs = exposed_specs(container)
    server = MCPServer("office", instructions=server_instructions(container))
    for spec in specs:
        register(server, spec)
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
        server = create_server(build_container(settings))
    except SettingsError as error:
        print(f"office-mcp: {error}", file=sys.stderr)
        raise SystemExit(2) from error
    serve(server, settings)


if __name__ == "__main__":
    main()
