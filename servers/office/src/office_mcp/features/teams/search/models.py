from __future__ import annotations

from dataclasses import asdict, dataclass
from enum import StrEnum
from typing import Any


class ResultType(StrEnum):
    MESSAGE = "message"
    GROUP_CHAT = "group_chat"
    PERSON = "person"
    FILE = "file"
    SUGGESTION = "suggestion"


@dataclass(frozen=True)
class SearchResult:
    type: ResultType
    text: str
    chat: str = ""
    members: str = ""
    time: str = ""

    def to_record(self) -> dict[str, Any]:
        return {**asdict(self), "type": self.type.value}
