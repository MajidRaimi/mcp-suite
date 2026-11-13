from __future__ import annotations

from dataclasses import dataclass

from whatsapp_mcp.features.attachments.service import AttachmentPolicy
from whatsapp_mcp.features.drafts.service import Draft, DraftStore
from whatsapp_mcp.features.waha.client import WahaClient
from whatsapp_mcp.shared.errors import PartialSendError, WhatsAppError


@dataclass(frozen=True)
class SentDraft:
    draft: Draft
    message_ids: list[str]


class DraftSender:
    def __init__(self, waha: WahaClient, drafts: DraftStore, attachments: AttachmentPolicy) -> None:
        self._waha = waha
        self._drafts = drafts
        self._attachments = attachments

    def send(self, draft_id: str, now: float) -> SentDraft:
        draft = self._drafts.take(draft_id, now)
        try:
            message_ids = self._deliver(draft)
        except PartialSendError:
            raise
        except WhatsAppError:
            self._drafts.restore(draft, now)
            raise
        return SentDraft(draft, message_ids)

    def _deliver(self, draft: Draft) -> list[str]:
        if not draft.attachments:
            return [self._waha.send_text(draft.chat_id, draft.text)]
        payloads = [(file, self._attachments.read_verified(file)) for file in draft.attachments]
        sent: list[str] = []
        for index, (file, data) in enumerate(payloads):
            caption = draft.text if index == 0 else None
            try:
                sent.append(
                    self._waha.send_file(draft.chat_id, file.filename, file.mimetype, data, caption)
                )
            except WhatsAppError as error:
                if sent:
                    raise PartialSendError(
                        f"sent {len(sent)} of {len(payloads)} items to {draft.chat_name}, then "
                        f"failed on {file.filename!r}: {error}. Do not resend the whole draft."
                    ) from error
                raise
        return sent
