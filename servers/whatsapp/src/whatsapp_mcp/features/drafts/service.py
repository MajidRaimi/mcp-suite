from __future__ import annotations

import fcntl
import json
import os
import secrets
import tempfile
import threading
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from pathlib import Path

from whatsapp_mcp.features.attachments.service import Attachment
from whatsapp_mcp.shared.errors import DraftError

_ID_ALPHABET = "abcdefghjkmnpqrstuvwxyz23456789"
_MAX_TEXT = 4_000


@dataclass(frozen=True)
class Draft:
    draft_id: str
    chat_id: str
    chat_name: str
    text: str
    created_at: float
    attachments: tuple[Attachment, ...] = ()


class DraftStore:
    def __init__(self, path: Path, ttl_seconds: int) -> None:
        self._path = path
        self._lock_path = path.with_name(f"{path.name}.lock")
        self._ttl = ttl_seconds
        self._thread_lock = threading.Lock()

    def create(
        self,
        chat_id: str,
        chat_name: str,
        text: str,
        now: float,
        attachments: tuple[Attachment, ...] = (),
    ) -> Draft:
        cleaned = text.strip()
        if not cleaned:
            raise DraftError("the reply text is empty")
        if len(cleaned) > _MAX_TEXT:
            raise DraftError(f"the reply is longer than {_MAX_TEXT} characters")
        draft = Draft(_new_id(), chat_id, chat_name, cleaned, now, attachments)
        with self._exclusive():
            drafts = self._live(now)
            drafts[draft.draft_id] = draft
            self._save(drafts)
        return draft

    def take(self, draft_id: str, now: float) -> Draft:
        with self._exclusive():
            drafts = self._live(now)
            draft = drafts.pop(draft_id.strip().lower(), None)
            if draft is None:
                raise DraftError(
                    f"draft {draft_id!r} does not exist or expired; create a new draft and show it "
                    "to the user before sending"
                )
            self._save(drafts)
        return draft

    def restore(self, draft: Draft, now: float) -> None:
        with self._exclusive():
            drafts = self._live(now)
            drafts[draft.draft_id] = draft
            self._save(drafts)

    @contextmanager
    def _exclusive(self) -> Iterator[None]:
        with self._thread_lock:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            descriptor = os.open(self._lock_path, os.O_RDWR | os.O_CREAT, 0o600)
            try:
                fcntl.flock(descriptor, fcntl.LOCK_EX)
                yield
            finally:
                os.close(descriptor)

    def _live(self, now: float) -> dict[str, Draft]:
        try:
            raw = json.loads(self._path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        drafts: dict[str, Draft] = {}
        if isinstance(raw, dict):
            for key, value in raw.items():
                if isinstance(value, dict):
                    draft = _parse_draft(value)
                    if draft is None:
                        continue
                    if now - draft.created_at <= self._ttl:
                        drafts[str(key)] = draft
        return drafts

    def _save(self, drafts: dict[str, Draft]) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({key: asdict(value) for key, value in drafts.items()})
        descriptor, temporary = tempfile.mkstemp(dir=self._path.parent, suffix=".tmp")
        try:
            with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
                handle.write(payload)
            os.chmod(temporary, 0o600)
            os.replace(temporary, self._path)
        except BaseException:
            Path(temporary).unlink(missing_ok=True)
            raise


def _parse_draft(value: dict[str, object]) -> Draft | None:
    draft_id, chat_id, chat_name, text = (
        value.get("draft_id"),
        value.get("chat_id"),
        value.get("chat_name"),
        value.get("text"),
    )
    created_at = value.get("created_at")
    raw_attachments = value.get("attachments", [])
    if not (
        isinstance(draft_id, str)
        and isinstance(chat_id, str)
        and isinstance(chat_name, str)
        and isinstance(text, str)
        and isinstance(created_at, int | float)
        and isinstance(raw_attachments, list)
    ):
        return None
    try:
        attachments = tuple(Attachment(**item) for item in raw_attachments)
    except TypeError:
        return None
    return Draft(draft_id, chat_id, chat_name, text, float(created_at), attachments)


def _new_id() -> str:
    return "".join(secrets.choice(_ID_ALPHABET) for _ in range(6))
