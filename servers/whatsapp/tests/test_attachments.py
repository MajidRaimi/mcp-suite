from __future__ import annotations

import base64
from pathlib import Path

import pytest
from whatsapp_fakes import NOW, FakeWaha, make_client

from whatsapp_mcp.features.attachments.service import AttachmentPolicy
from whatsapp_mcp.features.drafts.sender import DraftSender
from whatsapp_mcp.features.drafts.service import DraftStore
from whatsapp_mcp.shared.errors import (
    AttachmentError,
    DraftError,
    PartialSendError,
    WahaUnavailableError,
)


@pytest.fixture
def allowed(tmp_path: Path) -> Path:
    directory = tmp_path / "outbox"
    directory.mkdir()
    return directory


@pytest.fixture
def outside(tmp_path: Path) -> Path:
    secret = tmp_path / "secret.txt"
    secret.write_text("private")
    return secret


def _pdf(directory: Path, name: str = "proposal.pdf", body: bytes = b"%PDF-1.7 hello") -> Path:
    path = directory / name
    path.write_bytes(body)
    return path


def test_file_inside_the_allowed_directory_is_accepted(allowed: Path) -> None:
    pdf = _pdf(allowed)
    files = AttachmentPolicy(allowed).inspect_all([str(pdf)])
    assert files[0].mimetype == "application/pdf"
    assert files[0].path == str(pdf.resolve())


def test_file_outside_the_allowed_directory_is_rejected(allowed: Path, outside: Path) -> None:
    with pytest.raises(AttachmentError, match="outside the allowed"):
        AttachmentPolicy(allowed).inspect(str(outside))


def test_dot_dot_escape_is_rejected(allowed: Path, outside: Path) -> None:
    with pytest.raises(AttachmentError, match="outside the allowed"):
        AttachmentPolicy(allowed).inspect(f"{allowed}/../{outside.name}")


def test_symlink_escape_is_rejected(allowed: Path, outside: Path) -> None:
    link = allowed / "innocent.pdf"
    link.symlink_to(outside)
    with pytest.raises(AttachmentError, match="outside the allowed"):
        AttachmentPolicy(allowed).inspect(str(link))


def test_attachments_are_disabled_when_no_directory_is_configured(outside: Path) -> None:
    policy = AttachmentPolicy(None)
    assert policy.inspect_all([]) == ()
    with pytest.raises(AttachmentError, match="WHATSAPP_MCP_ATTACHMENTS_DIR"):
        policy.inspect_all([str(outside)])


def test_paths_must_be_absolute_existing_files(allowed: Path) -> None:
    policy = AttachmentPolicy(allowed)
    with pytest.raises(AttachmentError, match="absolute"):
        policy.inspect("relative.pdf")
    with pytest.raises(AttachmentError, match="does not exist"):
        policy.inspect(str(allowed / "missing.pdf"))
    with pytest.raises(AttachmentError, match="regular file"):
        policy.inspect(str(allowed))
    (allowed / "empty.pdf").write_bytes(b"")
    with pytest.raises(AttachmentError, match="empty"):
        policy.inspect(str(allowed / "empty.pdf"))


def test_send_time_check_rejects_a_file_swapped_for_a_symlink(allowed: Path, outside: Path) -> None:
    pdf = _pdf(allowed)
    policy = AttachmentPolicy(allowed)
    attachment = policy.inspect(str(pdf))
    pdf.unlink()
    pdf.symlink_to(outside)
    with pytest.raises(AttachmentError, match="outside the allowed"):
        policy.read_verified(attachment)


def _swap_after_confine(
    policy: AttachmentPolicy, monkeypatch: pytest.MonkeyPatch, victim: Path, target: Path
) -> None:
    confine = policy._confine

    def confine_then_swap(raw_path: str) -> Path:
        resolved = confine(raw_path)
        victim.rename(victim.with_name(f"{victim.name}.moved"))
        victim.symlink_to(target)
        return resolved

    monkeypatch.setattr(policy, "_confine", confine_then_swap)


def test_symlink_swapped_in_after_the_check_is_not_followed(
    allowed: Path, outside: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    pdf = _pdf(allowed)
    policy = AttachmentPolicy(allowed)
    _swap_after_confine(policy, monkeypatch, pdf, outside)
    with pytest.raises(AttachmentError, match="symlink"):
        policy.inspect(str(pdf))


def test_directory_swapped_for_a_symlink_after_the_check_is_not_followed(
    allowed: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    nested = allowed / "nested"
    nested.mkdir()
    pdf = _pdf(nested)
    elsewhere = tmp_path / "elsewhere"
    elsewhere.mkdir()
    _pdf(elsewhere, body=b"private")
    policy = AttachmentPolicy(allowed)
    attachment = policy.inspect(str(pdf))
    _swap_after_confine(policy, monkeypatch, nested, elsewhere)
    with pytest.raises(AttachmentError, match="symlink"):
        policy.read_verified(attachment)


def test_send_time_check_rejects_a_draft_made_under_a_wider_policy(
    allowed: Path, tmp_path: Path
) -> None:
    pdf = _pdf(tmp_path)
    attachment = AttachmentPolicy(tmp_path).inspect(str(pdf))
    with pytest.raises(AttachmentError, match="outside the allowed"):
        AttachmentPolicy(allowed).read_verified(attachment)


def test_send_time_check_detects_changed_content(allowed: Path) -> None:
    pdf = _pdf(allowed)
    policy = AttachmentPolicy(allowed)
    attachment = policy.inspect(str(pdf))
    pdf.write_bytes(b"%PDF-1.7 changed")
    with pytest.raises(AttachmentError, match="changed"):
        policy.read_verified(attachment)


def test_sender_sends_files_with_caption_on_the_first_one(tmp_path: Path, allowed: Path) -> None:
    waha = FakeWaha()
    policy = AttachmentPolicy(allowed)
    store = DraftStore(tmp_path / "drafts.json", ttl_seconds=600)
    files = policy.inspect_all([str(_pdf(allowed)), str(_pdf(allowed, "notes.txt", b"notes"))])
    draft = store.create("15550000001@c.us", "Ahmed", "the proposal", NOW, files)
    result = DraftSender(make_client(waha), store, policy).send(draft.draft_id, NOW)
    assert result.message_ids == ["file-1", "file-2"]
    first, second = waha.sent
    assert first["caption"] == "the proposal"
    assert first["file"] == {
        "mimetype": "application/pdf",
        "filename": "proposal.pdf",
        "data": base64.b64encode(b"%PDF-1.7 hello").decode("ascii"),
    }
    assert "caption" not in second
    with pytest.raises(DraftError):
        store.take(draft.draft_id, NOW)


def test_sender_restores_the_draft_when_nothing_was_sent(tmp_path: Path, allowed: Path) -> None:
    waha = FakeWaha()
    waha.fail_send = True
    store = DraftStore(tmp_path / "drafts.json", ttl_seconds=600)
    sender = DraftSender(make_client(waha), store, AttachmentPolicy(allowed))
    draft = store.create("15550000001@c.us", "Ahmed", "hello", NOW)
    with pytest.raises(WahaUnavailableError, match="500"):
        sender.send(draft.draft_id, NOW)
    waha.fail_send = False
    assert sender.send(draft.draft_id, NOW).message_ids == ["sent-1"]


def test_sender_reports_partial_multi_file_sends(tmp_path: Path, allowed: Path) -> None:
    waha = FakeWaha()
    waha.fail_file_after = 1
    policy = AttachmentPolicy(allowed)
    store = DraftStore(tmp_path / "drafts.json", ttl_seconds=600)
    files = policy.inspect_all([str(_pdf(allowed)), str(_pdf(allowed, "b.pdf"))])
    draft = store.create("15550000001@c.us", "Ahmed", "two files", NOW, files)
    with pytest.raises(PartialSendError, match="sent 1 of 2"):
        DraftSender(make_client(waha), store, policy).send(draft.draft_id, NOW)
    with pytest.raises(DraftError):
        store.take(draft.draft_id, NOW)
