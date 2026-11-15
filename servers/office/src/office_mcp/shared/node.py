from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass, field

STATIC_TEXT = "AXStaticText"


@dataclass(frozen=True, eq=False)
class Node:
    role: str
    title: str = ""
    value: str = ""
    description: str = ""
    dom_id: str = ""
    classes: tuple[str, ...] = ()
    children: tuple[Node, ...] = ()
    handle: object | None = field(default=None, repr=False)
    subrole: str = ""
    help: str = ""
    enabled: bool | None = None
    level: int | None = None
    selected: bool | None = None

    def walk(self) -> Iterator[Node]:
        yield self
        for child in self.children:
            yield from child.walk()

    def find(self, predicate: Callable[[Node], bool]) -> Node | None:
        return next((node for node in self.walk() if predicate(node)), None)

    def has_class(self, name: str) -> bool:
        return name in self.classes

    def text(self) -> str:
        pieces = (node.value for node in self.walk() if node.role == STATIC_TEXT)
        return " ".join(" ".join(piece.split()) for piece in pieces if piece.strip())

    def label(self) -> str:
        return " ".join((self.title or self.description or self.text() or self.help).split())
