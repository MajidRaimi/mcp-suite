from __future__ import annotations

import json
from datetime import UTC
from pathlib import Path

import httpx

from whatsapp_mcp.features.waha.client import WahaClient
from whatsapp_mcp.shared.settings import Settings

NOW = 1_800_000_000
API_KEY = "key"
SESSION = "default"


def message(
    timestamp: int, from_me: bool, body: str = "hi", participant: str | None = None
) -> dict[str, object]:
    return {
        "id": f"m{timestamp}",
        "timestamp": timestamp,
        "fromMe": from_me,
        "body": body,
        "hasMedia": False,
        "participant": participant,
    }


CHATS: list[dict[str, object]] = [
    {
        "id": "15550000001@c.us",
        "name": "Ahmed",
        "lastMessage": message(NOW - 3_600, False, "are you coming?"),
    },
    {"id": "15550000002@c.us", "name": "Sara", "lastMessage": message(NOW - 600, True, "done")},
    {
        "id": "1203630@g.us",
        "name": "Family",
        "lastMessage": message(NOW - 120, False, "dinner at 8", "15550000009@c.us"),
    },
    {
        "id": "15550000003@c.us",
        "name": "Old Friend",
        "lastMessage": message(NOW - 30 * 86_400, False, "long ago"),
    },
    {
        "id": "15550000004@c.us",
        "name": "Ahmed Work",
        "lastMessage": message(NOW - 7_200, False, "report?"),
    },
]
MESSAGES: dict[str, list[dict[str, object]]] = {
    "15550000001@c.us": [
        message(NOW - 3_600, False, "are you coming?"),
        message(NOW - 3_700, False, "hello"),
        message(NOW - 9_000, True, "see you"),
    ],
    "15550000004@c.us": [
        message(NOW - 7_200, False, "report?"),
        message(NOW - 8_000, True, "sure"),
    ],
    "1203630@g.us": [message(NOW - 120, False, "dinner at 8", "15550000009@c.us")],
    "15550000002@c.us": [
        message(NOW - 600, True, "done"),
        message(NOW - 700, False, "invoice attached"),
    ],
}


class FakeWaha:
    def __init__(self) -> None:
        self.sent: list[dict[str, object]] = []
        self.fail_send = False
        self.fail_file_after: int | None = None
        self.session_status = "WORKING"

    def __call__(self, request: httpx.Request) -> httpx.Response:
        if request.headers.get("X-Api-Key") != API_KEY:
            return httpx.Response(401, text="secret upstream detail")
        path = request.url.path
        if path == f"/api/{SESSION}/chats/overview":
            return httpx.Response(200, json=CHATS)
        if path.startswith(f"/api/{SESSION}/chats/") and path.endswith("/messages"):
            chat_id = path.split("/")[4]
            return httpx.Response(200, json=MESSAGES.get(chat_id, []))
        if path == "/api/sendText":
            if self.fail_send:
                return httpx.Response(500, text="engine down")
            self.sent.append(json.loads(request.content))
            return httpx.Response(201, json={"id": "sent-1"})
        if path == "/api/sendFile":
            if self.fail_file_after is not None and len(self.sent) >= self.fail_file_after:
                return httpx.Response(500, text="engine down")
            self.sent.append(json.loads(request.content))
            return httpx.Response(201, json={"id": f"file-{len(self.sent)}"})
        if path == f"/api/sessions/{SESSION}":
            return httpx.Response(200, json={"status": self.session_status})
        return httpx.Response(404)


def make_http(fake: FakeWaha) -> httpx.Client:
    return httpx.Client(transport=httpx.MockTransport(fake))


def make_client(fake: FakeWaha, api_key: str = API_KEY) -> WahaClient:
    return WahaClient("http://waha", api_key, SESSION, make_http(fake))


def make_settings(
    tmp_path: Path,
    *,
    waha_api_key: str = API_KEY,
    token: str | None = None,
    allowed_hosts: tuple[str, ...] = (),
    owner_name: str | None = None,
    attachments_dir: Path | None = None,
) -> Settings:
    return Settings(
        waha_url="http://waha",
        waha_api_key=waha_api_key,
        session=SESSION,
        host="127.0.0.1",
        port=8768,
        transport="streamable-http",
        token=token,
        allowed_hosts=allowed_hosts,
        timezone=UTC,
        owner_name=owner_name,
        attachments_dir=attachments_dir,
        drafts_path=tmp_path / "drafts.json",
        draft_ttl_minutes=30,
    )
