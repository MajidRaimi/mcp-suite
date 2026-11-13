from __future__ import annotations

import hmac

from mcp.server.transport_security import TransportSecuritySettings
from starlette.datastructures import Headers
from starlette.responses import JSONResponse
from starlette.types import ASGIApp, Receive, Scope, Send

LOOPBACK_HOSTS = ("127.0.0.1:*", "localhost:*", "[::1]:*")
BEARER = "bearer "


class BearerTokenMiddleware:
    def __init__(self, app: ASGIApp, token: str) -> None:
        self._app = app
        self._expected = token.encode()

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http" or self._authorized(Headers(scope=scope)):
            await self._app(scope, receive, send)
            return
        response = JSONResponse(
            {"error": "unauthorized"},
            status_code=401,
            headers={"WWW-Authenticate": "Bearer"},
        )
        await response(scope, receive, send)

    def _authorized(self, headers: Headers) -> bool:
        header = headers.get("authorization", "")
        if header[: len(BEARER)].lower() != BEARER:
            return False
        return hmac.compare_digest(header[len(BEARER) :].strip().encode(), self._expected)


def transport_security(allowed_hosts: tuple[str, ...]) -> TransportSecuritySettings | None:
    if not allowed_hosts:
        return None
    expanded = (pattern for host in allowed_hosts for pattern in host_patterns(host))
    hosts = list(dict.fromkeys((*LOOPBACK_HOSTS, *expanded)))
    origins = [f"{scheme}://{host}" for host in hosts for scheme in ("http", "https")]
    return TransportSecuritySettings(allowed_hosts=hosts, allowed_origins=origins)


def host_patterns(host: str) -> tuple[str, ...]:
    return (host,) if has_port(host) else (host, f"{host}:*")


def has_port(host: str) -> bool:
    return "]:" in host if host.startswith("[") else ":" in host
