from __future__ import annotations

import errno
import hashlib
import mimetypes
import os
import stat
from dataclasses import dataclass
from pathlib import Path

from whatsapp_mcp.shared.errors import AttachmentError

MAX_ATTACHMENT_MB = 64
MAX_ATTACHMENT_BYTES = MAX_ATTACHMENT_MB * 1024 * 1024
MAX_ATTACHMENTS = 10
ATTACHMENTS_DIR_VARIABLE = "WHATSAPP_MCP_ATTACHMENTS_DIR"


@dataclass(frozen=True)
class Attachment:
    path: str
    filename: str
    mimetype: str
    size: int
    sha256: str


class AttachmentPolicy:
    def __init__(self, allowed_dir: Path | None) -> None:
        self._allowed_dir = allowed_dir.resolve() if allowed_dir is not None else None

    def inspect_all(self, raw_paths: list[str]) -> tuple[Attachment, ...]:
        if not raw_paths:
            return ()
        if len(raw_paths) > MAX_ATTACHMENTS:
            raise AttachmentError(f"at most {MAX_ATTACHMENTS} attachments per draft")
        return tuple(self.inspect(raw_path) for raw_path in raw_paths)

    def inspect(self, raw_path: str) -> Attachment:
        path = self._confine(raw_path)
        data = self._read_confined(path, raw_path)
        if not data:
            raise AttachmentError(f"attachment is empty: {raw_path!r}")
        mimetype = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        return Attachment(str(path), path.name, mimetype, len(data), _digest(data))

    def read_verified(self, attachment: Attachment) -> bytes:
        path = self._confine(attachment.path)
        data = self._read_confined(path, attachment.filename)
        if _digest(data) != attachment.sha256:
            raise AttachmentError(
                f"attachment {attachment.filename!r} changed after the draft was shown; create a "
                "new draft"
            )
        return data

    def _confine(self, raw_path: str) -> Path:
        root = self._require_root()
        path = Path(raw_path).expanduser()
        if not path.is_absolute():
            raise AttachmentError(f"attachment path must be absolute: {raw_path!r}")
        try:
            resolved = path.resolve(strict=True)
        except OSError as error:
            raise AttachmentError(f"attachment does not exist: {raw_path!r}") from error
        if not resolved.is_relative_to(root):
            raise AttachmentError(
                f"attachment {raw_path!r} is outside the allowed attachments directory "
                f"{str(root)!r}"
            )
        return resolved

    def _read_confined(self, resolved: Path, label: str) -> bytes:
        root = self._require_root()
        relative = resolved.relative_to(root)
        if not relative.parts:
            raise AttachmentError(f"attachment is not a regular file: {label!r}")
        try:
            descriptor = _open_beneath(root, relative)
        except OSError as error:
            raise _open_error(error, label) from error
        with os.fdopen(descriptor, "rb") as handle:
            status = os.fstat(handle.fileno())
            if not stat.S_ISREG(status.st_mode):
                raise AttachmentError(f"attachment is not a regular file: {label!r}")
            _ensure_within_limit(status.st_size, resolved.name)
            try:
                data = handle.read(MAX_ATTACHMENT_BYTES + 1)
            except OSError as error:
                raise _unreadable(error, label) from error
        _ensure_within_limit(len(data), resolved.name)
        return data

    def _require_root(self) -> Path:
        if self._allowed_dir is None:
            raise AttachmentError(
                f"file attachments are disabled: set {ATTACHMENTS_DIR_VARIABLE} to the directory "
                "that files may be attached from, then restart the server"
            )
        return self._allowed_dir


def _ensure_within_limit(size: int, filename: str) -> None:
    if size > MAX_ATTACHMENT_BYTES:
        raise AttachmentError(f"attachment {filename!r} is larger than {MAX_ATTACHMENT_MB} MB")


def _open_beneath(root: Path, relative: Path) -> int:
    *directories, filename = relative.parts
    parent = os.open(root, os.O_RDONLY | os.O_DIRECTORY)
    try:
        for directory in directories:
            child = os.open(directory, os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW, dir_fd=parent)
            os.close(parent)
            parent = child
        return os.open(filename, os.O_RDONLY | os.O_NOFOLLOW | os.O_NONBLOCK, dir_fd=parent)
    finally:
        os.close(parent)


def _open_error(error: OSError, label: str) -> AttachmentError:
    if error.errno in (errno.ELOOP, errno.ENOTDIR):
        return AttachmentError(
            f"attachment {label!r} changed into a symlink while it was being read; create a new "
            "draft"
        )
    if error.errno == errno.EISDIR:
        return AttachmentError(f"attachment is not a regular file: {label!r}")
    return _unreadable(error, label)


def _unreadable(error: OSError, label: str) -> AttachmentError:
    return AttachmentError(f"attachment {label!r} is unreadable: {error.strerror}")


def _digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()
