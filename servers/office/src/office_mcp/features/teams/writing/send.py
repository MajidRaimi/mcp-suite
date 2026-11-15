from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from office_mcp.features.teams.header import open_thread_id
from office_mcp.features.teams.messages.models import Message
from office_mcp.features.teams.writing.compose import Commit, ComposeBox, choose_commit, strip_quote
from office_mcp.features.teams.writing.context import WriteContext
from office_mcp.features.teams.writing.guard import TeamsSendGuard
from office_mcp.features.teams.writing.reply import Quoter
from office_mcp.features.teams.writing.targets import (
    OpenedTarget,
    require_meeting_ok,
    require_message_id,
    require_size_ok,
    require_text,
)
from office_mcp.features.teams.writing.verify import (
    find_sent,
    newer_own,
    newest_from_others,
    newest_id,
    recent_duplicate,
    sent_early,
)
from office_mcp.shared.drafts import Channel, Draft, Review, TextGuard
from office_mcp.shared.errors import (
    DuplicateSendError,
    OfficeError,
    SendRefusedError,
    UiTimeoutError,
    error_text,
)
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.node import Node
from office_mcp.shared.results import SendResult, SendStatus
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import UiScope

VERIFY_SECONDS = 10.0
MIN_VERIFY_SECONDS = 2.0
EARLY_SEND = (
    "Teams sent part of the text early, so Send was not pressed; check the chat before "
    "doing anything else"
)
DRY_RUN_NOTE = "dry run: every check passed; nothing was typed or sent"
DISCARDED_NOTE = "the text typed into the compose box was removed again"
KEPT_NOTE = "the compose box could not be cleared; check it before sending again"
SWITCHED_NOTE = (
    "Teams switched to another chat before the compose box was cleared, so nothing was "
    "deleted; check the compose box of the original chat"
)
EARLY_KEPT_NOTE = "the rest of the text may still be in the compose box"
OUT_OF_TIME = "the time for this send ran out before Send was pressed; nothing was sent"
STILL_IN_BOX = "the text is still in the compose box, so Teams may not have sent it"


@dataclass(frozen=True)
class SendRequest:
    chat: str
    text: str
    group_ok: bool = False
    allow_language_mismatch: bool = False
    allow_duplicate: bool = False
    dry_run: bool = False
    reply_to: str | None = None


@dataclass(frozen=True)
class Prepared:
    target: OpenedTarget
    review: Review
    quoted: Message | None

    @property
    def text(self) -> str:
        return self.review.draft.body

    @property
    def thread_id(self) -> str:
        return self.target.thread_id


@dataclass(frozen=True)
class Delivery:
    status: SendStatus
    message_ref: str
    verified: bool
    notes: tuple[str, ...]


DRY_RUN = Delivery(status=SendStatus.DRY_RUN, message_ref="", verified=False, notes=(DRY_RUN_NOTE,))


@dataclass(frozen=True)
class SendOutcome:
    prepared: Prepared
    delivery: Delivery

    def result(self, focus: FocusOutcome, notes: tuple[str, ...]) -> SendResult:
        target = self.prepared.target
        extra: dict[str, Any] = {
            "chat": target.name,
            "thread_id": target.thread_id,
            "members": target.members,
        }
        quoted = self.prepared.quoted
        if quoted is not None:
            extra["quoted"] = {"message_id": quoted.id, "author": quoted.author}
        return SendResult(
            status=self.delivery.status,
            channel=Channel.TEAMS,
            targets=(target.thread_id,),
            subject="",
            message_ref=self.delivery.message_ref,
            verified=self.delivery.verified,
            review=self.prepared.review,
            focus=focus,
            notes=(*self.delivery.notes, *notes),
            extra=extra,
        )


def first_draft(request: SendRequest) -> Draft:
    return Draft(
        channel=Channel.TEAMS,
        body=request.text,
        allow_language_mismatch=request.allow_language_mismatch,
    )


def context_draft(request: SendRequest, target: OpenedTarget) -> Draft:
    context = newest_from_others(target.messages)
    return Draft(
        channel=Channel.TEAMS,
        body=request.text,
        context=context.text if context is not None else "",
        context_author=context.author if context is not None else "",
        recipients=(target.thread_id,),
        allow_language_mismatch=request.allow_language_mismatch,
    )


class TeamsSender:
    def __init__(
        self,
        context: WriteContext,
        box: ComposeBox,
        quoter: Quoter,
        guard: TextGuard,
        duplicates: TeamsSendGuard,
        max_members: int,
    ) -> None:
        self._context = context
        self._box = box
        self._quoter = quoter
        self._guard = guard
        self._duplicates = duplicates
        self._max_members = max_members

    def send(self, request: SendRequest, deadline: Deadline) -> SendResult:
        require_text(request.text)
        if request.reply_to is not None:
            require_message_id(request.reply_to)
        self._guard.require_clean(first_draft(request))
        scoped = self._context.run(lambda scope: self._run(scope, request, deadline))
        return scoped.value.result(scoped.focus, scoped.notes)

    def _run(self, scope: UiScope, request: SendRequest, deadline: Deadline) -> SendOutcome:
        prepared = self._prepare(request)
        if request.dry_run:
            return SendOutcome(prepared, DRY_RUN)
        scope.ensure_front()
        return SendOutcome(prepared, self._deliver(prepared, deadline))

    def _prepare(self, request: SendRequest) -> Prepared:
        targets = self._context.targets
        row = targets.resolve(request.chat)
        self._context.recipients.require_thread(row.chat.thread_id)
        require_meeting_ok(row, request.group_ok)
        opened = targets.open(row)
        require_size_ok(row, opened.members, request.group_ok, self._max_members)
        quoted: Message | None = None
        if request.reply_to is not None:
            opened, quoted = targets.locate(opened, request.reply_to.strip())
        self._box.require_empty(self._box.locate(opened.root))
        review = self._guard.require_clean(context_draft(request, opened))
        require_text(review.draft.body)
        if not request.allow_duplicate:
            self._refuse_duplicate(opened, review.draft.body)
        return Prepared(target=opened, review=review, quoted=quoted)

    def _refuse_duplicate(self, target: OpenedTarget, text: str) -> None:
        self._duplicates.check(target.thread_id, text)
        duplicate = recent_duplicate(target.messages, text)
        if duplicate is not None:
            raise DuplicateSendError(
                f"your recent message {duplicate.id} in this chat already says the same; "
                "nothing was sent. Pass allow_duplicate=true only when the owner wants it again."
            )

    def _deliver(self, prepared: Prepared, deadline: Deadline) -> Delivery:
        targets = self._context.targets
        root = targets.current(prepared.thread_id)
        box = self._box.locate(root)
        self._box.require_empty(box)
        before = newest_id(targets.messages(root))
        quote = ""
        try:
            if prepared.quoted is not None:
                quote = self._quoter.add_quote(prepared.thread_id, prepared.quoted).in_box
            commit = self._insert(prepared, quote, before, deadline)
        except Exception as error:
            note = self._clean_up(error, prepared.thread_id, box, quote, prepared.text)
            if note:
                error.add_note(note)
            raise
        return self._commit(prepared, box, quote, commit, before, deadline)

    def _insert(self, prepared: Prepared, quote: str, before: int, deadline: Deadline) -> Commit:
        box = self._box.locate(self._context.targets.now(prepared.thread_id))
        self._box.write(box, prepared.text)
        root = self._unsent(prepared.thread_id, before)
        shown = strip_quote(self._box.read(self._box.locate(root)), quote)
        self._guard.verify_written(prepared.review.draft, shown, "")
        root = self._unsent(prepared.thread_id, before)
        if deadline.expired():
            raise UiTimeoutError(OUT_OF_TIME)
        commit = choose_commit(root)
        self._box.require_commit_target(self._box.locate(root), commit)
        return commit

    def _clean_up(self, error: Exception, thread_id: str, box: Node, quote: str, text: str) -> str:
        if isinstance(error, SendRefusedError):
            return EARLY_KEPT_NOTE
        if not self._still_open(thread_id):
            return KEPT_NOTE
        if self._box.discard(box, quote, text, lambda: self._still_open(thread_id)):
            return DISCARDED_NOTE
        if not self._still_open(thread_id):
            return SWITCHED_NOTE
        return "" if self._box.is_clear(box) else KEPT_NOTE

    def _still_open(self, thread_id: str) -> bool:
        try:
            self._context.targets.now(thread_id)
        except OfficeError:
            return False
        return True

    def _unsent(self, thread_id: str, before: int) -> Node:
        targets = self._context.targets
        root = targets.now(thread_id)
        if sent_early(targets.messages(root), before):
            raise SendRefusedError(EARLY_SEND)
        return root

    def _commit(
        self,
        prepared: Prepared,
        box: Node,
        quote: str,
        commit: Commit,
        before: int,
        deadline: Deadline,
    ) -> Delivery:
        notes: list[str] = []
        try:
            self._box.commit(commit)
        except Exception as error:
            notes.append(f"the {commit.path} reported an error: {describe(error)}")
        self._duplicates.record(prepared.thread_id, prepared.text)
        sent = self._observe(prepared, before, deadline, notes)
        if sent is not None:
            return Delivery(SendStatus.SENT, sent.id, True, tuple(notes))
        if self._box.still_holds(box, quote, prepared.text):
            notes.append(STILL_IN_BOX)
        return Delivery(SendStatus.UNCONFIRMED, "", False, tuple(notes))

    def _observe(
        self, prepared: Prepared, before: int, deadline: Deadline, notes: list[str]
    ) -> Message | None:
        timeout = min(VERIFY_SECONDS, max(deadline.remaining(), MIN_VERIFY_SECONDS))
        quoted = prepared.quoted is not None
        try:
            sent = self._context.waiter.until(lambda: self._sent(prepared, before, quoted), timeout)
            if sent is None:
                notes.append(self._missing(prepared, before, timeout))
            return sent
        except Exception as error:
            notes.append(f"checking the chat after Send failed: {describe(error)}")
            return None

    def _sent(self, prepared: Prepared, before: int, quoted: bool) -> Message | None:
        root = self._context.gateway.snapshot()
        if open_thread_id(root) != prepared.thread_id:
            return None
        messages = self._context.targets.messages(root)
        return find_sent(messages, before, prepared.text, quoted)

    def _missing(self, prepared: Prepared, before: int, timeout: float) -> str:
        root = self._context.gateway.snapshot()
        others = newer_own(self._context.targets.messages(root), before)
        if others:
            ids = ", ".join(message.id for message in others)
            return f"a new message ({ids}) appeared, but its text differs from the text sent"
        return f"Teams did not show the new message within {timeout:.0f} s"


def describe(error: Exception) -> str:
    return error_text(error) if isinstance(error, OfficeError) else repr(error)
