from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class Message(BaseModel):
    model_config = ConfigDict(extra="ignore", populate_by_name=True)

    id: str = ""
    timestamp: int = 0
    from_me: bool = Field(default=False, alias="fromMe")
    body: str | None = None
    has_media: bool = Field(default=False, alias="hasMedia")
    participant: str | None = None
    author: str | None = None


class ChatSummary(BaseModel):
    model_config = ConfigDict(extra="ignore", populate_by_name=True)

    id: str
    name: str | None = None
    last_message: Message | None = Field(default=None, alias="lastMessage")

    @property
    def is_group(self) -> bool:
        return self.id.endswith("@g.us")

    @property
    def display_name(self) -> str:
        return self.name or self.id.split("@")[0]
