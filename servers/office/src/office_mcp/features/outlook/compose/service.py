from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field, replace
from datetime import datetime
from typing import Any

from office_mcp.features.outlook.compose.gateway import (
    ComposeGateway,
    ComposeSurface,
    RecipientField,
    ReplyKind,
)
from office_mcp.features.outlook.compose.guard import SendGuard, SentRecord, send_fingerprint
from office_mcp.features.outlook.compose.recipients import (
    ExpectedRecipients,
    reply_recipients,
    require_distinct,
    require_tokens,
    token_addresses,
    validate_addresses,
)
from office_mcp.features.outlook.compose.sensitivity import validate_sensitivity
from office_mcp.features.outlook.compose.session import (
    ComposeSession,
    Observation,
    run_scoped,
)
from office_mcp.features.outlook.compose.undo import UndoResult, UndoSend
from office_mcp.features.outlook.compose.verify import (
    TOP_ROWS,
    display_subject,
    new_matches,
    recent_duplicates,
    require_subject,
)
from office_mcp.features.outlook.conversation import ConversationReader, LatestMessage
from office_mcp.features.outlook.models import Row
from office_mcp.features.outlook.rows import KEYS_ONLY, FolderLocator, RowFinder, is_key
from office_mcp.shared.drafts import Channel, Draft, Review, TextGuard, canonical_text
from office_mcp.shared.errors import (
    ComposeError,
    DuplicateSendError,
    OfficeError,
    SendRefusedError,
    TargetError,
    TargetNotFoundError,
    TextEntryError,
    UiTimeoutError,
    error_text,
)
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.results import SendResult, SendStatus
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import ScopeFactory, ScopeOptions, UiScope
from office_mcp.shared.waiting import Waiter

SEND_SCOPE = ScopeOptions(foreground=True)
SENT_FOLDERS = ("Sent Items", "Sent")
OUTBOX_FOLDERS = ("Outbox",)
MAX_SUBJECT = 255
MAX_BODY = 20000
VERIFY_SECONDS = 10.0
CLOSE_SECONDS = 10.0
MIN_SECONDS_TO_SEND = 3.0
DRY_RUN_NOTE = "dry run: nothing was opened or sent"
STILL_OPEN_NOTE = "the compose window was still open after Send"


@dataclass(frozen=True)
class NewMail:
    to: tuple[str, ...]
    subject: str
    body: str
    cc: tuple[str, ...] = ()
    bcc: tuple[str, ...] = ()
    allow_language_mismatch: bool = False
    allow_duplicate: bool = False
    dry_run: bool = False
    sensitivity: tuple[str, ...] = ()


@dataclass(frozen=True)
class ReplyMail:
    key: str
    body: str
    reply_all: bool = False
    cc: tuple[str, ...] = ()
    bcc: tuple[str, ...] = ()
    folder: str | None = None
    allow_language_mismatch: bool = False
    allow_duplicate: bool = False
    dry_run: bool = False
    sensitivity: tuple[str, ...] = ()


@dataclass(frozen=True)
class ForwardMail:
    key: str
    to: tuple[str, ...]
    body: str = ""
    cc: tuple[str, ...] = ()
    bcc: tuple[str, ...] = ()
    folder: str | None = None
    allow_duplicate: bool = False
    dry_run: bool = False
    sensitivity: tuple[str, ...] = ()


@dataclass(frozen=True)
class SendPlan:
    review: Review
    expected: ExpectedRecipients
    typed: ExpectedRecipients
    subject: str
    fingerprint: str
    allow_duplicate: bool
    dry_run: bool
    kind: ReplyKind | None = None
    key: str = ""
    original: LatestMessage | None = None
    folder: str | None = None
    sensitivity: tuple[str, ...] = ()
    extra: dict[str, Any] = field(default_factory=dict)

    @property
    def match_subject(self) -> str:
        if self.kind is None or self.original is None:
            return self.subject
        return display_subject(self.kind, self.original.subject)


@dataclass(frozen=True)
class KnownRows:
    sent: frozenset[str]
    outbox: frozenset[str]


@dataclass(frozen=True)
class Delivery:
    status: SendStatus
    subject: str
    verified: bool
    message_ref: str
    notes: tuple[str, ...]


@dataclass(frozen=True)
class ComposeDeps:
    gateway: ComposeGateway
    rows: RowFinder
    scope: ScopeFactory
    reader: ConversationReader
    guard: TextGuard
    recipients: RecipientGuard
    sends: SendGuard
    waiter: Waiter
    now: Callable[[], datetime]
    owner_email: Callable[[], str | None]


class ComposeService:
    def __init__(self, deps: ComposeDeps) -> None:
        self._deps = deps
        self._undo = UndoSend(
            deps.gateway, deps.scope, deps.sends, deps.waiter, self._outgoing_rows, deps.now
        )

    def send(self, mail: NewMail, deadline: Deadline, fail_at: str | None = None) -> SendResult:
        return self._deliver(self._plan_new(mail), deadline, fail_at)

    def reply(self, mail: ReplyMail, deadline: Deadline, fail_at: str | None = None) -> SendResult:
        return self._deliver(self._plan_reply(mail), deadline, fail_at)

    def forward(
        self, mail: ForwardMail, deadline: Deadline, fail_at: str | None = None
    ) -> SendResult:
        return self._deliver(self._plan_forward(mail), deadline, fail_at)

    def undo_send(self) -> UndoResult:
        return self._undo.undo()

    def _plan_new(self, mail: NewMail) -> SendPlan:
        to = validate_addresses(RecipientField.TO, mail.to, 1)
        cc = validate_addresses(RecipientField.CC, mail.cc, 0)
        bcc = validate_addresses(RecipientField.BCC, mail.bcc, 0)
        require_distinct((*to, *cc, *bcc))
        subject = validate_subject(mail.subject)
        body = validate_body(mail.body, required=True)
        expected = ExpectedRecipients(to=to, cc=cc, bcc=bcc)
        review = self._deps.guard.require_clean(
            Draft(
                channel=Channel.EMAIL,
                body=body,
                subject=subject,
                sender_account=self._deps.owner_email() or "",
                recipients=expected.everyone(),
                allow_language_mismatch=mail.allow_language_mismatch,
            )
        )
        return SendPlan(
            review=review,
            expected=expected,
            typed=expected,
            subject=review.draft.subject,
            fingerprint=send_fingerprint(
                "send", expected.everyone(), review.draft.subject, review.draft.body
            ),
            allow_duplicate=mail.allow_duplicate,
            sensitivity=mail.sensitivity,
            dry_run=mail.dry_run,
        )

    def _plan_reply(self, mail: ReplyMail) -> SendPlan:
        require_key(mail.key)
        cc = validate_addresses(RecipientField.CC, mail.cc, 0)
        bcc = validate_addresses(RecipientField.BCC, mail.bcc, 0)
        body = validate_body(mail.body, required=True)
        latest = self._deps.reader.latest_message(mail.key, mail.folder)
        owner = self._deps.owner_email()
        prefilled = reply_recipients(latest, owner, mail.reply_all)
        expected = ExpectedRecipients(prefilled, cc, bcc, merge_to_cc=mail.reply_all)
        require_distinct(expected.everyone())
        review = self._deps.guard.require_clean(
            Draft(
                channel=Channel.EMAIL,
                body=body,
                context=latest.body_top,
                context_author=latest.sender.name if latest.sender is not None else "",
                sender_account=owner or "",
                recipients=expected.everyone(),
                allow_language_mismatch=mail.allow_language_mismatch,
            )
        )
        kind = ReplyKind.REPLY_ALL if mail.reply_all else ReplyKind.REPLY
        return SendPlan(
            review=review,
            expected=expected,
            typed=ExpectedRecipients((), cc, bcc),
            subject="",
            fingerprint=send_fingerprint(
                f"{kind.value}:{mail.key}", expected.everyone(), "", review.draft.body
            ),
            allow_duplicate=mail.allow_duplicate,
            sensitivity=mail.sensitivity,
            dry_run=mail.dry_run,
            kind=kind,
            key=mail.key,
            original=latest,
            folder=mail.folder or latest.folder or None,
            extra={
                "in_reply_to": {"key": mail.key, "subject": latest.subject},
                "recipients": {
                    "to": list(expected.to),
                    "cc": list(expected.cc),
                    "bcc": list(expected.bcc),
                },
            },
        )

    def _plan_forward(self, mail: ForwardMail) -> SendPlan:
        require_key(mail.key)
        to = validate_addresses(RecipientField.TO, mail.to, 1)
        cc = validate_addresses(RecipientField.CC, mail.cc, 0)
        bcc = validate_addresses(RecipientField.BCC, mail.bcc, 0)
        require_distinct((*to, *cc, *bcc))
        body = validate_body(mail.body, required=False)
        latest = self._deps.reader.latest_message(mail.key, mail.folder)
        expected = ExpectedRecipients(to=to, cc=cc, bcc=bcc)
        review = self._forward_review(body, expected)
        return SendPlan(
            review=review,
            expected=expected,
            typed=expected,
            subject="",
            fingerprint=send_fingerprint(
                f"forward:{mail.key}", expected.everyone(), "", review.draft.body
            ),
            allow_duplicate=mail.allow_duplicate,
            sensitivity=mail.sensitivity,
            dry_run=mail.dry_run,
            kind=ReplyKind.FORWARD,
            key=mail.key,
            original=latest,
            folder=mail.folder or latest.folder or None,
            extra={"forwarded": {"key": mail.key, "subject": latest.subject}},
        )

    def _forward_review(self, body: str, expected: ExpectedRecipients) -> Review:
        draft = Draft(
            channel=Channel.EMAIL,
            body=body,
            sender_account=self._deps.owner_email() or "",
            recipients=expected.everyone(),
        )
        if body.strip():
            return self._deps.guard.require_clean(draft)
        version = self._deps.guard.review(draft).persona_version
        return Review(draft=draft, violations=(), fixes=(), persona_version=version)

    def _deliver(self, plan: SendPlan, deadline: Deadline, fail_at: str | None) -> SendResult:
        plan = labelled(plan)
        self._deps.recipients.require_emails(plan.expected.everyone())
        if not plan.allow_duplicate:
            self._deps.sends.check(plan.fingerprint)
        if plan.dry_run:
            return self._result(plan, dry_run_delivery(plan), FocusOutcome.NOT_NEEDED, ())
        scoped = run_scoped(
            self._deps.scope,
            SEND_SCOPE,
            lambda ui: self._send_in_scope(ui, plan, deadline, fail_at),
        )
        return self._result(plan, scoped.value, scoped.focus, scoped.notes)

    def _send_in_scope(
        self, ui: UiScope, plan: SendPlan, deadline: Deadline, fail_at: str | None
    ) -> Delivery:
        known = self._known_rows(plan)
        opener = self._opener(plan)
        with ComposeSession(
            self._deps.gateway,
            opener,
            self._deps.waiter,
            fail_at,
            keep_busy=plan.kind is not None,
        ) as session:
            session.require_empty_body()
            self._fill(session, plan)
            session.set_sensitivity(plan.sensitivity)
            subject = self._check(session, plan)
            require_time(deadline)
            ui.ensure_front()
            entry = self._commit(session, plan, subject)
            try:
                observation = self._observe(session, deadline)
            except SendRefusedError:
                self._deps.sends.forget(entry)
                raise
        return self._confirm(subject, observation, known, deadline)

    def _commit(self, session: ComposeSession, plan: SendPlan, subject: str) -> SentRecord:
        try:
            session.commit()
        except BaseException:
            if session.committed:
                self._deps.sends.record(plan.fingerprint, subject)
            raise
        return self._deps.sends.record(plan.fingerprint, subject)

    def _opener(self, plan: SendPlan) -> Callable[[], ComposeSurface]:
        gateway = self._deps.gateway
        if plan.kind is None:
            return gateway.open_new
        rows = self._deps.rows
        row = rows.refresh(plan.key, rows.locator_for(plan.key, plan.folder))
        kind = plan.kind
        return lambda: gateway.open_reply(row.node, kind)

    def _fill(self, session: ComposeSession, plan: SendPlan) -> None:
        typed = plan.typed
        session.fill(
            (
                (RecipientField.TO, typed.to),
                (RecipientField.CC, typed.cc),
                (RecipientField.BCC, typed.bcc),
            ),
            plan.review.draft.subject if plan.kind is None else "",
            plan.review.draft.body,
        )

    def _check(self, session: ComposeSession, plan: SendPlan) -> str:
        body = session.written_body()
        subject = session.subject()
        if plan.kind is None:
            self._deps.guard.verify_written(plan.review.draft, body, subject)
        else:
            original = plan.original.subject if plan.original is not None else ""
            require_subject(plan.kind, original, subject)
            if plan.review.draft.body:
                self._deps.guard.verify_written(plan.review.draft, body, "")
            elif canonical_text(body):
                raise TextEntryError("the forward body should be empty but holds text")
        self._check_recipients(session, plan.expected)
        session.require_sensitivity(plan.sensitivity)
        return subject

    def _check_recipients(self, session: ComposeSession, expected: ExpectedRecipients) -> None:
        to = token_addresses(RecipientField.TO, session.recipients(RecipientField.TO))
        cc = token_addresses(RecipientField.CC, session.recipients(RecipientField.CC))
        bcc = token_addresses(RecipientField.BCC, session.recipients(RecipientField.BCC))
        if expected.merge_to_cc:
            require_tokens("To and Cc", (*to, *cc), (*expected.to, *expected.cc))
        else:
            require_tokens("To", to, expected.to)
            require_tokens("Cc", cc, expected.cc)
        require_tokens("Bcc", bcc, expected.bcc)
        self._deps.recipients.require_emails((*to, *cc, *bcc))

    def _observe(self, session: ComposeSession, deadline: Deadline) -> Observation:
        try:
            return session.observe(min(CLOSE_SECONDS, deadline.remaining()))
        except SendRefusedError:
            raise
        except OfficeError as error:
            return Observation(
                closed=False,
                notes=(f"could not watch the compose after Send: {error_text(error)}",),
            )

    def _known_rows(self, plan: SendPlan) -> KnownRows:
        sent = self._top_rows(SENT_FOLDERS)
        outbox = self._top_rows(OUTBOX_FOLDERS)
        if not plan.allow_duplicate:
            self._refuse_duplicates(plan.match_subject, sent, outbox)
        return KnownRows(
            sent=frozenset(row.summary.key for row in sent),
            outbox=frozenset(row.summary.key for row in outbox),
        )

    def _refuse_duplicates(self, subject: str, sent: list[Row], outbox: list[Row]) -> None:
        now = self._deps.now()
        if recent_duplicates(sent, subject, now) or recent_duplicates(outbox, subject, now):
            raise DuplicateSendError(
                f"a mail with the subject {subject[:120]!r} was sent in the last 10 minutes "
                "(Sent Items or Outbox); pass allow_duplicate=true only if it must go out again"
            )

    def _outgoing_rows(self) -> list[Row]:
        outbox = self._top_rows(OUTBOX_FOLDERS)[:TOP_ROWS]
        return [*outbox, *self._top_rows(SENT_FOLDERS)[:TOP_ROWS]]

    def _top_rows(self, names: tuple[str, ...]) -> list[Row]:
        for name in names:
            try:
                return self._deps.rows.rows(FolderLocator(name))
            except TargetNotFoundError:
                continue
        return []

    def _confirm(
        self, subject: str, observation: Observation, known: KnownRows, deadline: Deadline
    ) -> Delivery:
        notes = observation.notes
        if not observation.closed:
            return Delivery(SendStatus.UNCONFIRMED, subject, False, "", (*notes, STILL_OPEN_NOTE))
        try:
            sent = self._deps.waiter.until(
                lambda: first_new(self._top_rows(SENT_FOLDERS), subject, known.sent),
                min(VERIFY_SECONDS, deadline.remaining()),
            )
            if sent is not None:
                return Delivery(SendStatus.SENT, subject, True, sent.summary.key, notes)
            queued = first_new(self._top_rows(OUTBOX_FOLDERS), subject, known.outbox)
            if queued is not None:
                return Delivery(SendStatus.SUBMITTED, subject, False, queued.summary.key, notes)
        except OfficeError as error:
            notes = (*notes, f"could not check Sent Items: {error_text(error)}")
        return Delivery(SendStatus.UNCONFIRMED, subject, False, "", notes)

    def _result(
        self, plan: SendPlan, delivery: Delivery, focus: FocusOutcome, notes: tuple[str, ...]
    ) -> SendResult:
        return SendResult(
            status=delivery.status,
            channel=Channel.EMAIL,
            targets=plan.expected.everyone(),
            subject=delivery.subject,
            message_ref=delivery.message_ref,
            verified=delivery.verified,
            review=plan.review,
            focus=focus,
            notes=(*delivery.notes, *notes),
            extra=plan.extra,
        )


def labelled(plan: SendPlan) -> SendPlan:
    path = validate_sensitivity(plan.sensitivity)
    if not path:
        return plan
    return replace(plan, sensitivity=path, extra={**plan.extra, "sensitivity": list(path)})


def dry_run_delivery(plan: SendPlan) -> Delivery:
    return Delivery(SendStatus.DRY_RUN, plan.match_subject, False, "", (DRY_RUN_NOTE,))


def first_new(rows: list[Row], subject: str, known: frozenset[str]) -> Row | None:
    found = new_matches(rows, subject, known)
    return found[0] if found else None


def require_key(key: str) -> None:
    if not is_key(key):
        raise TargetError(KEYS_ONLY)


def require_time(deadline: Deadline) -> None:
    if deadline.remaining() < MIN_SECONDS_TO_SEND:
        raise UiTimeoutError("the send ran out of time, so it was stopped before pressing Send")


def validate_subject(subject: str) -> str:
    text = subject.strip()
    if not text or len(text) > MAX_SUBJECT:
        raise ComposeError(f"subject takes 1 to {MAX_SUBJECT} characters")
    if "\n" in text or "\r" in text:
        raise ComposeError("subject must be a single line")
    return text


def validate_body(body: str, required: bool) -> str:
    if len(body) > MAX_BODY:
        raise ComposeError(f"body takes at most {MAX_BODY} characters")
    if required and not body.strip():
        raise ComposeError("body must not be empty")
    return body
