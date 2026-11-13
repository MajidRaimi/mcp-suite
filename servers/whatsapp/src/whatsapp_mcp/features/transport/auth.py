from __future__ import annotations

import hmac

from starlette.datastructures import Headers
from starlette.responses import JSONResponse
from starlette.types import ASGIApp, Receive, Scope, Send

_SCHEME = "bearer"


class BearerTokenMiddleware:
    def __init__(self, app: ASGIApp, token: str) -> None:
        if not token:
            raise ValueError("BearerTokenMiddleware needs a non-empty token")
        self._app = app
        self._expected = token.encode("utf-8")

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
        scheme, _, credentials = headers.get("authorization", "").partition(" ")
        if scheme.lower() != _SCHEME:
            return False
        return hmac.compare_digest(credentials.strip().encode("utf-8"), self._expected)
