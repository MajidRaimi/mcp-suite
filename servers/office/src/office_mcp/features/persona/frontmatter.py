from __future__ import annotations

import re
import tomllib
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from office_mcp.shared.errors import PersonaFormatError

FENCE = "+++"
FIRST_TOML_LINE = 2
POSITION = re.compile(r"\s*\(at line (?P<line>\d+), column \d+\)$|\s*\(at end of document\)$")


@dataclass(frozen=True)
class Document:
    front_matter: str | None
    prose: str


def split_document(text: str, path: Path) -> Document:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines or lines[0].strip() != FENCE:
        return Document(front_matter=None, prose=text.strip())
    for index in range(1, len(lines)):
        if lines[index].strip() == FENCE:
            return Document(
                front_matter="\n".join(lines[1:index]),
                prose="\n".join(lines[index + 1 :]).strip(),
            )
    raise PersonaFormatError(
        f"{path}: line 1: the front matter opened with {FENCE} has no closing {FENCE} line"
    )


def parse_toml(front_matter: str, path: Path) -> dict[str, Any]:
    try:
        return tomllib.loads(front_matter)
    except tomllib.TOMLDecodeError as error:
        raise PersonaFormatError(toml_error_text(str(error), front_matter, path)) from error


def toml_error_text(message: str, front_matter: str, path: Path) -> str:
    found = POSITION.search(message)
    reason = message[: found.start()] if found else message
    if found and found.group("line"):
        toml_line = int(found.group("line"))
    else:
        toml_line = max(1, front_matter.count("\n") + 1)
    line = toml_line + FIRST_TOML_LINE - 1
    return f"{path}: line {line}: the front matter is not valid TOML: {reason}"
