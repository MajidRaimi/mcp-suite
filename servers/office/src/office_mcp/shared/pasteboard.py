from __future__ import annotations

from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Protocol

PLAIN_TEXT_TYPE = "public.utf8-plain-text"
TRANSIENT_TYPE = "org.nspasteboard.TransientType"
CONCEALED_TYPE = "org.nspasteboard.ConcealedType"


@dataclass(frozen=True)
class PasteboardSnapshot:
    items: tuple[Mapping[str, bytes], ...]


CLEARED = PasteboardSnapshot(())


class Pasteboard(Protocol):
    def snapshot(self) -> PasteboardSnapshot: ...

    def restore(self, snapshot: PasteboardSnapshot) -> None: ...

    def write_transient_text(self, text: str) -> None: ...


@contextmanager
def preserved(pasteboard: Pasteboard) -> Iterator[None]:
    saved = pasteboard.snapshot()
    try:
        yield
    finally:
        pasteboard.restore(saved)
