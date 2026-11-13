from __future__ import annotations

import time
from collections.abc import Callable, Sequence
from dataclasses import asdict
from typing import Annotated

import anyio
import httpx
import uvicorn
from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations
from pydantic import Field
from starlette.requests import Request
from starlette.responses import JSONResponse

from whatsapp_mcp import descriptions
from whatsapp_mcp.features.attachments.service import Attachment, AttachmentPolicy
from whatsapp_mcp.features.drafts.sender import DraftSender
from whatsapp_mcp.features.drafts.service import DraftStore
from whatsapp_mcp.features.inbox.service import InboxService
from whatsapp_mcp.features.transport.cli import parse_cli
from whatsapp_mcp.features.transport.http import build_http_app
from whatsapp_mcp.features.waha.client import WahaClient
from whatsapp_mcp.shared.errors import WhatsAppError
from whatsapp_mcp.shared.owner import Owner
from whatsapp_mcp.shared.settings import Settings

Result = dict[str, object]

SEARCH_HIT_LIMIT = 60
HEALTHY_STATUS = "WORKING"
UNREACHABLE_STATUS = "UNREACHABLE"

READ_ONLY = ToolAnnotations(
    read_only_hint=True, destructive_hint=False, idempotent_hint=True, open_world_hint=True
)
DRAFT = ToolAnnotations(
    read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=False
)
SEND = ToolAnnotations(
    read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=True
)


def build_server(settings: Settings, http: httpx.Client) -> MCPServer:
    owner = Owner(settings.owner_name)
    waha = WahaClient(settings.waha_url, settings.waha_api_key, settings.session, http)
    inbox = InboxService(waha, settings.timezone)
    attachment_policy = AttachmentPolicy(settings.attachments_dir)
    drafts = DraftStore(settings.drafts_path, settings.draft_ttl_minutes * 60)
    sender = DraftSender(waha, drafts, attachment_policy)
    server = MCPServer(name="whatsapp", instructions=descriptions.instructions(owner))

    @server.tool(
        name="whatsapp_needs_reply",
        description=descriptions.needs_reply(owner),
        annotations=READ_ONLY,
    )
    def needs_reply(
        limit: Annotated[int, Field(ge=1, le=50)] = 20,
        include_groups: bool = False,
        max_age_days: Annotated[int, Field(ge=1, le=60)] = 7,
    ) -> Result:
        return _guard(
            lambda: {
                "waiting": [
                    asdict(item)
                    for item in inbox.needs_reply(time.time(), limit, include_groups, max_age_days)
                ]
            }
        )

    @server.tool(
        name="whatsapp_recent_chats",
        description=descriptions.RECENT_CHATS,
        annotations=READ_ONLY,
    )
    def recent_chats(limit: Annotated[int, Field(ge=1, le=100)] = 20) -> Result:
        return _guard(lambda: {"chats": [asdict(item) for item in inbox.recent_chats(limit)]})

    @server.tool(
        name="whatsapp_read_chat",
        description=descriptions.READ_CHAT,
        annotations=READ_ONLY,
    )
    def read_chat(chat: str, limit: Annotated[int, Field(ge=1, le=200)] = 40) -> Result:
        def load() -> Result:
            summary, messages = inbox.read_chat(chat, limit)
            return {
                "chat_id": summary.id,
                "name": summary.display_name,
                "is_group": summary.is_group,
                "messages": [asdict(message) for message in messages],
            }

        return _guard(load)

    @server.tool(
        name="whatsapp_search",
        description=descriptions.SEARCH,
        annotations=READ_ONLY,
    )
    def search(
        query: Annotated[str, Field(min_length=2)],
        days: Annotated[int, Field(ge=1, le=60)] = 14,
        chats: Annotated[int, Field(ge=1, le=80)] = 40,
    ) -> Result:
        return _guard(
            lambda: {
                "hits": [asdict(hit) for hit in inbox.search(query, time.time(), days, chats)][
                    :SEARCH_HIT_LIMIT
                ]
            }
        )

    @server.tool(
        name="whatsapp_draft_reply",
        description=descriptions.draft_reply(owner),
        annotations=DRAFT,
    )
    def draft_reply(
        chat: str,
        text: Annotated[str, Field(min_length=1)],
        attachments: list[str] | None = None,
    ) -> Result:
        def create() -> Result:
            files = attachment_policy.inspect_all(attachments or [])
            summary = inbox.resolve(chat)
            draft = drafts.create(summary.id, summary.display_name, text, time.time(), files)
            return {
                "draft_id": draft.draft_id,
                "to": draft.chat_name,
                "chat_id": draft.chat_id,
                "text": draft.text,
                "attachments": [_describe(file) for file in draft.attachments],
                "expires_in_minutes": settings.draft_ttl_minutes,
                "status": "draft only, nothing was sent",
            }

        return _guard(create)

    @server.tool(
        name="whatsapp_send_draft",
        description=descriptions.send_draft(owner),
        annotations=SEND,
    )
    def send_draft(draft_id: str) -> Result:
        def send() -> Result:
            sent = sender.send(draft_id, time.time())
            return {
                "sent": True,
                "to": sent.draft.chat_name,
                "text": sent.draft.text,
                "attachments": [file.filename for file in sent.draft.attachments],
                "message_ids": sent.message_ids,
            }

        return _guard(send)

    @server.tool(
        name="whatsapp_status",
        description=descriptions.STATUS,
        annotations=READ_ONLY,
    )
    def status() -> Result:
        return _guard(lambda: {"session": settings.session, "status": waha.session_status()})

    async def health(request: Request) -> JSONResponse:
        state = await anyio.to_thread.run_sync(lambda: _health_status(waha))
        return JSONResponse({"status": state}, status_code=200 if state == HEALTHY_STATUS else 503)

    server.custom_route("/health", methods=["GET"], include_in_schema=False)(health)
    return server


def _health_status(waha: WahaClient) -> str:
    try:
        return waha.session_status()
    except WhatsAppError:
        return UNREACHABLE_STATUS


def _describe(file: Attachment) -> Result:
    return {"filename": file.filename, "mimetype": file.mimetype, "size_bytes": file.size}


def _guard(action: Callable[[], Result]) -> Result:
    try:
        return action()
    except WhatsAppError as error:
        return {"error": str(error)}


def main(argv: Sequence[str] | None = None) -> None:
    options = parse_cli(argv)
    try:
        settings = Settings.from_env().with_overrides(options.transport, options.host, options.port)
    except WhatsAppError as error:
        raise SystemExit(f"whatsapp-mcp: {error}") from error
    with httpx.Client() as http:
        server = build_server(settings, http)
        if settings.transport == "stdio":
            server.run("stdio")
            return
        uvicorn.run(build_http_app(server, settings), host=settings.host, port=settings.port)


if __name__ == "__main__":
    main()
