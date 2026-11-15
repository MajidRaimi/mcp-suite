from __future__ import annotations

from collections.abc import Mapping, Sequence

from office_mcp.shared.pasteboard import (
    CONCEALED_TYPE,
    PLAIN_TEXT_TYPE,
    TRANSIENT_TYPE,
    PasteboardSnapshot,
)


class FakePasteboard:
    def __init__(self, items: Sequence[Mapping[str, bytes]] | None = None) -> None:
        self.items: list[dict[str, bytes]] = [dict(item) for item in items or []]
        self.writes: list[str] = []
        self.restores = 0

    def snapshot(self) -> PasteboardSnapshot:
        return PasteboardSnapshot(tuple(dict(item) for item in self.items))

    def restore(self, snapshot: PasteboardSnapshot) -> None:
        self.restores += 1
        self.items = [dict(item) for item in snapshot.items]

    def write_transient_text(self, text: str) -> None:
        self.writes.append(text)
        self.items = [{PLAIN_TEXT_TYPE: text.encode(), TRANSIENT_TYPE: b"", CONCEALED_TYPE: b""}]

    def text(self) -> str | None:
        if not self.items or PLAIN_TEXT_TYPE not in self.items[0]:
            return None
        return self.items[0][PLAIN_TEXT_TYPE].decode()
