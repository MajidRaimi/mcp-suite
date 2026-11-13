from __future__ import annotations

import base64

import httpx
from pydantic import TypeAdapter, ValidationError

from whatsapp_mcp.features.waha.models import ChatSummary, Message
from whatsapp_mcp.shared.errors import WahaUnavailableError

_CHATS = TypeAdapter(list[ChatSummary])
_MESSAGES = TypeAdapter(list[Message])


class WahaClient:
    def __init__(self, base_url: str, api_key: str, session: str, http: httpx.Client) -> None:
        self._base_url = base_url
        self._session = session
        self._http = http
        self._headers = {"X-Api-Key": api_key}

    def session_status(self) -> str:
        body = self._get(f"/api/sessions/{self._session}", {})
        status = body.get("status") if isinstance(body, dict) else None
        return str(status or "UNKNOWN")

    def chats(self, limit: int, offset: int = 0) -> list[ChatSummary]:
        body = self._get(f"/api/{self._session}/chats/overview", {"limit": limit, "offset": offset})
        return self._parse(_CHATS, body)

    def messages(self, chat_id: str, limit: int) -> list[Message]:
        body = self._get(
            f"/api/{self._session}/chats/{chat_id}/messages",
            {"limit": limit, "downloadMedia": "false", "sortBy": "timestamp", "sortOrder": "desc"},
        )
        return self._parse(_MESSAGES, body)

    def send_text(self, chat_id: str, text: str) -> str:
        return self._post("/api/sendText", {"chatId": chat_id, "text": text})

    def send_file(
        self, chat_id: str, filename: str, mimetype: str, data: bytes, caption: str | None
    ) -> str:
        payload: dict[str, object] = {
            "chatId": chat_id,
            "file": {
                "mimetype": mimetype,
                "filename": filename,
                "data": base64.b64encode(data).decode("ascii"),
            },
        }
        if caption:
            payload["caption"] = caption
        return self._post("/api/sendFile", payload, timeout=120)

    def _post(self, path: str, payload: dict[str, object], timeout: float = 30) -> str:
        try:
            response = self._http.post(
                f"{self._base_url}{path}",
                json={"session": self._session, **payload},
                headers=self._headers,
                timeout=timeout,
            )
        except httpx.HTTPError as error:
            raise WahaUnavailableError(f"WAHA send failed: {error}") from error
        if response.status_code >= 400:
            raise WahaUnavailableError(
                f"WAHA send returned {response.status_code}: {response.text[:200]}"
            )
        body = response.json()
        identifier = body.get("id") if isinstance(body, dict) else None
        return str(identifier or "")

    def _get(self, path: str, params: dict[str, str | int]) -> object:
        try:
            response = self._http.get(
                f"{self._base_url}{path}", params=params, headers=self._headers, timeout=30
            )
        except httpx.HTTPError as error:
            raise WahaUnavailableError(f"WAHA is unreachable: {error}") from error
        if response.status_code >= 400:
            raise WahaUnavailableError(
                f"WAHA {path} returned {response.status_code}: {response.text[:200]}"
            )
        return response.json()

    @staticmethod
    def _parse[T](adapter: TypeAdapter[list[T]], body: object) -> list[T]:
        try:
            return adapter.validate_python(body)
        except ValidationError as error:
            raise WahaUnavailableError(f"unexpected WAHA response: {error}") from error
