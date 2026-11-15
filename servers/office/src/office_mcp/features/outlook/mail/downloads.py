from __future__ import annotations

import re
import shutil
from collections.abc import Callable
from pathlib import Path
from typing import Protocol

from office_mcp.shared.errors import AttachmentError, UiTimeoutError

PARTIAL_SUFFIXES = (".download", ".crdownload", ".part", ".partial")
CONTROL_CHARACTERS = re.compile(r"[\x00-\x1f\x7f]")
SEPARATORS = re.compile(r"[/\\:]")
SHOWN_SIZE = re.compile(
    r"^(?P<amount>\d+(?:[.,]\d+)?)\s*(?P<unit>bytes|b|kb|mb|gb)$", re.IGNORECASE
)
UNIT_POWERS = {"bytes": 0, "b": 0, "kb": 1, "mb": 2, "gb": 3}
SIZE_TOLERANCE = 0.1
FALLBACK_NAME = "attachment"
DIRECTORY_MODE = 0o700
STABLE_SECONDS = 1.0
POLL_SECONDS = 0.25


class DownloadsWatcher(Protocol):
    def snapshot(self) -> frozenset[str]: ...

    def wait_for(self, name: str, before: frozenset[str], timeout: float) -> Path: ...


def safe_name(name: str) -> str:
    base = name.replace("\\", "/").rsplit("/", 1)[-1]
    cleaned = CONTROL_CHARACTERS.sub("", SEPARATORS.sub("", base)).strip()
    return cleaned.lstrip(".").strip() or FALLBACK_NAME


def numbered(name: str, number: int) -> str:
    path = Path(name)
    return f"{path.stem} ({number}){path.suffix}"


def unique_path(directory: Path, name: str) -> Path:
    candidate = directory / name
    number = 1
    while candidate.exists():
        candidate = directory / numbered(name, number)
        number += 1
    return candidate


def download_pattern(name: str) -> re.Pattern[str]:
    path = Path(safe_name(name))
    return re.compile(rf"^{re.escape(path.stem)}(?: \(\d+\))?{re.escape(path.suffix)}$")


def is_partial(name: str) -> bool:
    return name.casefold().endswith(PARTIAL_SUFFIXES)


def size_matches(shown: str, actual: int) -> bool:
    match = SHOWN_SIZE.match(shown.strip())
    if match is None:
        return actual > 0
    amount = float(match["amount"].replace(",", "."))
    power = UNIT_POWERS[match["unit"].casefold()]
    for base in (1000, 1024):
        expected = amount * base**power
        slack = max(expected * SIZE_TOLERANCE, float(base**power))
        if abs(actual - expected) <= slack:
            return True
    return False


def move_into(source: Path, directory: Path, name: str) -> Path:
    directory.mkdir(mode=DIRECTORY_MODE, parents=True, exist_ok=True)
    target = unique_path(directory, safe_name(name))
    if target.resolve().parent != directory.resolve():
        raise AttachmentError(f"refused to save {name!r} outside {directory}")
    shutil.move(source, target)
    return target


class PathDownloadsWatcher:
    def __init__(
        self,
        directory: Path,
        clock: Callable[[], float],
        sleep: Callable[[float], None],
        stable_seconds: float = STABLE_SECONDS,
        poll: float = POLL_SECONDS,
    ) -> None:
        self._directory = directory
        self._clock = clock
        self._sleep = sleep
        self._stable_seconds = stable_seconds
        self._poll = poll

    def snapshot(self) -> frozenset[str]:
        if not self._directory.is_dir():
            return frozenset()
        return frozenset(entry.name for entry in self._directory.iterdir())

    def wait_for(self, name: str, before: frozenset[str], timeout: float) -> Path:
        pattern = download_pattern(name)
        deadline = self._clock() + timeout
        last: tuple[Path, int, float] | None = None
        while self._clock() < deadline:
            found = self._candidate(pattern, before)
            if found is not None:
                size = found.stat().st_size
                if last is None or last[0] != found or last[1] != size:
                    last = (found, size, self._clock())
                elif size > 0 and self._clock() - last[2] >= self._stable_seconds:
                    return found
            self._sleep(self._poll)
        raise UiTimeoutError(
            f"the attachment {name!r} did not finish downloading in {timeout:.0f} s"
        )

    def _candidate(self, pattern: re.Pattern[str], before: frozenset[str]) -> Path | None:
        names = self.snapshot()
        partial_stems = {entry[: entry.rfind(".")] for entry in names if is_partial(entry)}
        for entry in sorted(names - before):
            if is_partial(entry) or entry in partial_stems or not pattern.match(entry):
                continue
            path = self._directory / entry
            if path.is_file():
                return path
        return None
