from __future__ import annotations

import argparse
from collections.abc import Sequence
from dataclasses import dataclass

from whatsapp_mcp.shared.errors import ConfigurationError
from whatsapp_mcp.shared.settings import TRANSPORTS, Transport, parse_port, parse_transport


@dataclass(frozen=True)
class CliOptions:
    transport: Transport | None = None
    host: str | None = None
    port: int | None = None


def parse_cli(argv: Sequence[str] | None = None) -> CliOptions:
    namespace = _parser().parse_args(argv)
    return CliOptions(
        transport=_typed_transport(namespace.transport),
        host=_typed_text(namespace.host),
        port=_typed_port(namespace.port),
    )


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="whatsapp-mcp",
        description="MCP server for WhatsApp through a self-hosted WAHA.",
    )
    parser.add_argument(
        "--transport",
        choices=TRANSPORTS,
        help="stdio or streamable-http (default: WHATSAPP_MCP_TRANSPORT or streamable-http)",
    )
    parser.add_argument(
        "--host", help="HTTP bind address (default: WHATSAPP_MCP_HOST or 127.0.0.1)"
    )
    parser.add_argument(
        "--port",
        type=_port_argument,
        help="HTTP port (default: WHATSAPP_MCP_PORT or 8768)",
    )
    return parser


def _port_argument(raw: str) -> int:
    try:
        return parse_port(raw, "--port")
    except ConfigurationError as error:
        raise argparse.ArgumentTypeError(str(error)) from error


def _typed_transport(value: object) -> Transport | None:
    return parse_transport(value, "--transport") if isinstance(value, str) else None


def _typed_text(value: object) -> str | None:
    return value if isinstance(value, str) and value else None


def _typed_port(value: object) -> int | None:
    return value if isinstance(value, int) else None
