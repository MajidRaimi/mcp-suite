from __future__ import annotations

import hashlib
import stat
import threading
from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

from office_mcp.features.persona.frontmatter import parse_toml, split_document
from office_mcp.features.persona.models import (
    EMPTY_PERSONA,
    Persona,
    PersonaState,
    PersonaStatus,
    Rules,
)
from office_mcp.features.persona.schema import parse_rules
from office_mcp.shared.drafts import NO_PERSONA_VERSION
from office_mcp.shared.errors import PersonaError, PersonaFormatError, PersonaMissingError

VERSION_HEX_CHARS = 12
PRIVATE_BITS = stat.S_IRWXG | stat.S_IRWXO
NANOSECONDS = 1_000_000_000


@dataclass(frozen=True)
class FileStamp:
    mtime_ns: int
    size: int
    mode: int


@dataclass(frozen=True)
class PersonaSnapshot:
    status: PersonaStatus
    persona: Persona | None

    def require(self) -> Persona:
        if self.persona is not None:
            return self.persona
        error = self.status.error or f"the persona file is {self.status.state.value}"
        if self.status.state is PersonaState.MISSING:
            raise PersonaMissingError(error)
        raise PersonaFormatError(error)


class PersonaStore:
    def __init__(
        self, path: Path, explicit: bool, required: bool, clock: Callable[[], datetime]
    ) -> None:
        self._path = path
        self._explicit = explicit
        self._required = required
        self._clock = clock
        self._lock = threading.Lock()
        self._cached: tuple[FileStamp, PersonaSnapshot] | None = None

    @property
    def path(self) -> Path:
        return self._path

    @property
    def required(self) -> bool:
        return self._required

    def snapshot(self) -> PersonaSnapshot:
        with self._lock:
            try:
                stamp = self._stamp()
            except OSError as error:
                self._cached = None
                return self._unreadable(error, None, True)
            if stamp is None:
                self._cached = None
                return self._absent()
            if self._cached is not None and self._cached[0] == stamp:
                return self._cached[1]
            snapshot = self._load(stamp)
            self._cached = (stamp, snapshot)
            return snapshot

    def current(self) -> Persona:
        return self.snapshot().require()

    def status(self) -> PersonaStatus:
        return self.snapshot().status

    @property
    def version(self) -> str:
        return self.status().version

    def _stamp(self) -> FileStamp | None:
        try:
            info = self._path.stat()
        except FileNotFoundError:
            return None
        return FileStamp(mtime_ns=info.st_mtime_ns, size=info.st_size, mode=info.st_mode)

    def _absent(self) -> PersonaSnapshot:
        if self._explicit or self._required:
            reason = "OFFICE_MCP_PERSONA_FILE" if self._explicit else "OFFICE_MCP_PERSONA_REQUIRED"
            status = self._status(
                PersonaState.MISSING,
                error=f"the persona file {self._path} does not exist ({reason} is set)",
            )
            return PersonaSnapshot(status=status, persona=None)
        return PersonaSnapshot(status=self._status(PersonaState.NONE), persona=EMPTY_PERSONA)

    def _load(self, stamp: FileStamp) -> PersonaSnapshot:
        modified_at = datetime.fromtimestamp(stamp.mtime_ns / NANOSECONDS, tz=UTC)
        permissions_ok = not stamp.mode & PRIVATE_BITS
        try:
            content = self._path.read_bytes()
        except OSError as error:
            return self._unreadable(error, modified_at, permissions_ok)
        version = content_version(content)
        try:
            persona = parse_persona(content, self._path)
        except PersonaError as error:
            status = self._status(
                PersonaState.INVALID,
                version=version,
                error=str(error),
                modified_at=modified_at,
                permissions_ok=permissions_ok,
            )
            return PersonaSnapshot(status=status, persona=None)
        loaded = Persona(
            prose=persona.prose,
            rules=persona.rules,
            version=version,
            source=self._path,
            modified_at=modified_at,
            loaded_at=self._clock(),
        )
        status = self._status(
            PersonaState.LOADED,
            version=version,
            modified_at=modified_at,
            permissions_ok=permissions_ok,
        )
        return PersonaSnapshot(status=status, persona=loaded)

    def _unreadable(
        self, error: OSError, modified_at: datetime | None, permissions_ok: bool
    ) -> PersonaSnapshot:
        status = self._status(
            PersonaState.INVALID,
            error=f"{self._path}: the persona file cannot be read: {error.strerror or error}",
            modified_at=modified_at,
            permissions_ok=permissions_ok,
        )
        return PersonaSnapshot(status=status, persona=None)

    def _status(
        self,
        state: PersonaState,
        version: str = NO_PERSONA_VERSION,
        error: str | None = None,
        modified_at: datetime | None = None,
        permissions_ok: bool = True,
    ) -> PersonaStatus:
        return PersonaStatus(
            state=state,
            path=self._path,
            version=version,
            modified_at=modified_at,
            permissions_ok=permissions_ok,
            error=error,
        )


def parse_persona(content: bytes, path: Path) -> Persona:
    try:
        text = content.decode("utf-8")
    except UnicodeDecodeError as error:
        raise PersonaFormatError(f"{path}: the persona file is not valid UTF-8") from error
    document = split_document(text, path)
    if document.front_matter is None:
        return Persona(prose=document.prose, rules=Rules())
    data = parse_toml(document.front_matter, path)
    return Persona(prose=document.prose, rules=parse_rules(data, document.front_matter, path))


def content_version(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()[:VERSION_HEX_CHARS]
