from __future__ import annotations

import hmac

from mcp.server.transport_security import TransportSecuritySettings
from starlette.datastructures import Headers
from starlette.responses import JSONResponse
from starlette.types import ASGIApp, Receive, Scope, Send

LOOPBACK_HOSTS = ("127.0.0.1:*", "localhost:*", "[::1]:*")
ORIGIN_SCHEMES = ("http", "https")


class BearerTokenMiddleware:
    def __init__(self, app: ASGIApp, token: str) -> None:
        self._app = app
        self._expected = f"Bearer {token}".encode()

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http" or self._is_authorized(scope):
            await self._app(scope, receive, send)
            return
        response = JSONResponse(
            {"error": "unauthorized"},
            status_code=401,
            headers={"WWW-Authenticate": "Bearer"},
        )
        await response(scope, receive, send)

    def _is_authorized(self, scope: Scope) -> bool:
        provided = Headers(scope=scope).get("authorization", "").encode()
        return hmac.compare_digest(provided, self._expected)


def transport_security(allowed_hosts: tuple[str, ...]) -> TransportSecuritySettings | None:
    if not allowed_hosts:
        return None
    hosts = [*LOOPBACK_HOSTS, *(pattern for host in allowed_hosts for pattern in _patterns(host))]
    return TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=hosts,
        allowed_origins=[f"{scheme}://{host}" for host in hosts for scheme in ORIGIN_SCHEMES],
    )


def _patterns(host: str) -> tuple[str, ...]:
    has_port = ":" in host[host.rfind("]") + 1 :]
    return (host,) if has_port else (host, f"{host}:*")
