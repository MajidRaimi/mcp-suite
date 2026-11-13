from __future__ import annotations

from dataclasses import dataclass

DEFAULT_OWNER_LABEL = "the user"


@dataclass(frozen=True)
class Owner:
    label: str

    @classmethod
    def from_name(cls, name: str | None) -> Owner:
        return cls(label=name or DEFAULT_OWNER_LABEL)

    @property
    def possessive(self) -> str:
        return f"{self.label}'s"
