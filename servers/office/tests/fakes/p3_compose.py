from __future__ import annotations

from collections.abc import Callable, Sequence
from contextlib import AbstractContextManager
from dataclasses import dataclass, field

from fakes.clock import FakeClock, fake_waiter
from fakes.drivers import FAKE_NOW, FAKE_TODAY
from fakes.guard import RecordingGuard
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.scope import RecordingScope, recording_factory
from fakes.trees import folder_row, outlook_row
from office_mcp.features.outlook.compose.gateway import (
    AlertSheet,
    ComposeSurface,
    RecipientField,
    RecipientRead,
    RecipientToken,
    ReplyKind,
)
from office_mcp.features.outlook.compose.guard import SendGuard
from office_mcp.features.outlook.compose.service import ComposeDeps, ComposeService
from office_mcp.features.outlook.conversation import LatestMessage
from office_mcp.features.outlook.listing import parse_rows
from office_mcp.features.outlook.models import Person
from office_mcp.features.outlook.rows import RefRegistry, RowFinder
from office_mcp.shared.errors import CleanupError, OfficeError
from office_mcp.shared.node import Node
from office_mcp.shared.recipients import RecipientGuard, parse_allowlist
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import ScopeFactory, ScopeOptions, UiScope

OWNER = "jane@example.com"
SIGNATURE = "\nJane Doe\nExample Corp"
HISTORY = "\nJane Doe\n\nFrom: Omar Haddad <omar@example.com>\nSent: Monday\nSubject: x"
SENT_FOLDER = "Sent Items"
INBOX_ROW = "Omar Haddad, Project Atlas budget,     9:15 AM,        Numbers attached"


def no_change(text: str) -> str:
    return text


@dataclass
class FakeCompose:
    baseline: str = SIGNATURE
    reply_baseline: str = HISTORY
    reply_subject: str = "RE: Project Atlas budget"
    prefilled: dict[RecipientField, list[str]] = field(default_factory=dict)
    extra_tokens: dict[RecipientField, list[RecipientToken]] = field(default_factory=dict)
    pending: dict[RecipientField, str] = field(default_factory=dict)
    unreadable: set[str] = field(default_factory=set)
    body_transform: Callable[[str], str] = no_change
    alert_after_send: AlertSheet | None = None
    close_on_send: bool = True
    on_send: Callable[[], None] | None = None
    open_error: OfficeError | None = None
    stray_on_error: bool = False
    stray_subject: str = ""
    undo_enabled: bool = False
    reopen_subject: str | None = None
    send_error: Exception | None = None
    discard_error: OfficeError | None = None
    shown_as: dict[str, str] = field(default_factory=dict)
    label_shown_as: dict[str, str] = field(default_factory=dict)
    sheets: list[tuple[object, AlertSheet]] = field(default_factory=list)
    calls: list[str] = field(default_factory=list)
    open: list[ComposeSurface] = field(default_factory=list)
    send_presses: int = 0
    discards: int = 0
    _body: str = ""
    _base: str = ""
    _subject: str = ""
    _label: str = ""
    _tokens: dict[RecipientField, list[RecipientToken]] = field(default_factory=dict)
    _count: int = 0

    def surfaces(self) -> list[ComposeSurface]:
        return list(self.open)

    def stray_surfaces(self, before: Sequence[ComposeSurface]) -> list[ComposeSurface]:
        return [surface for surface in self.open if all(surface is not b for b in before)]

    def open_new(self) -> ComposeSurface:
        self.calls.append("open_new")
        if self.open_error is not None:
            if self.stray_on_error:
                self._new_surface("", self.stray_subject, {})
            raise self.open_error
        return self._new_surface(self.baseline, "", {})

    def open_reply(self, row: Node, kind: ReplyKind) -> ComposeSurface:
        self.calls.append(f"open_reply:{row.handle}:{kind.value}")
        return self._new_surface(self.reply_baseline, self.reply_subject, self.prefilled)

    def dismiss_popovers(self) -> int:
        self.calls.append("dismiss_popovers")
        return 0

    def body(self, surface: ComposeSurface) -> str:
        return self._body + self._base

    def insert_body(self, surface: ComposeSurface, text: str) -> None:
        self.calls.append("insert_body")
        self._body = self.body_transform(text) + self._body

    def subject(self, surface: ComposeSurface) -> str:
        return self._subject

    def set_subject(self, surface: ComposeSurface, text: str) -> None:
        self.calls.append("set_subject")
        self._subject = text

    def add_recipient(self, surface: ComposeSurface, field: RecipientField, address: str) -> None:
        self.calls.append(f"add:{field.value}:{address}")
        self._tokens.setdefault(field, []).append(self._token(address))

    def recipients(self, surface: ComposeSurface, field: RecipientField) -> RecipientRead:
        tokens = (*self._tokens.get(field, []), *self.extra_tokens.get(field, []))
        return RecipientRead(tokens=tokens, pending=self.pending.get(field, ""))

    def set_sensitivity(self, surface: ComposeSurface, path: tuple[str, ...]) -> None:
        self.calls.append(f"sensitivity:{'/'.join(path)}")
        self._label = self.label_shown_as.get(path[-1], path[-1])

    def shown_texts(self, surface: ComposeSurface) -> list[str]:
        return [self._label] if self._label else []

    def send_button(self, surface: ComposeSurface) -> Node:
        return Node(role="AXButton", title="Send", handle="send")

    def press_send(self, button: Node) -> None:
        self.calls.append("press_send")
        self.send_presses += 1
        if self.alert_after_send is not None and self.open:
            self.sheets.append((self.open[-1].window, self.alert_after_send))
            return
        if self.close_on_send:
            self.open.clear()
            if self.on_send is not None:
                self.on_send()
        if self.send_error is not None:
            raise self.send_error

    def is_open(self, surface: ComposeSurface) -> bool:
        return any(surface is candidate for candidate in self.open)

    def alerts(self, window: object) -> list[AlertSheet]:
        return [sheet for owner, sheet in self.sheets if owner == window]

    def new_alert(self, window: object, before: Sequence[AlertSheet]) -> AlertSheet | None:
        fresh = [sheet for sheet in self.alerts(window) if all(sheet is not b for b in before)]
        return fresh[0] if fresh else None

    def press_alert(self, alert: AlertSheet, title: str) -> None:
        self.calls.append(f"alert:{title}")
        self.sheets = [entry for entry in self.sheets if entry[1] is not alert]

    def discard(self, surface: ComposeSurface) -> None:
        self.calls.append("discard")
        if self.discard_error is not None:
            raise self.discard_error
        self.discards += 1
        self.open = [candidate for candidate in self.open if candidate is not surface]

    def undo_send_enabled(self) -> bool:
        return self.undo_enabled

    def undo_send(self) -> None:
        self.calls.append("undo_send")
        if self.reopen_subject is not None:
            self._new_surface(self.baseline, self.reopen_subject, {})

    def _new_surface(
        self, base: str, subject: str, prefilled: dict[RecipientField, list[str]]
    ) -> ComposeSurface:
        self._count += 1
        name = f"compose{self._count}"
        surface = ComposeSurface(handle=name, window=name, inline=False)
        self.open.append(surface)
        self._base = base
        self._body = ""
        self._subject = subject
        self._label = ""
        self._tokens = {key: [self._token(a) for a in value] for key, value in prefilled.items()}
        return surface

    def _token(self, address: str) -> RecipientToken:
        if address in self.unreadable:
            return RecipientToken(label="Somebody", address=None)
        shown = self.shown_as.get(address, address)
        return RecipientToken(label=f"Person, {shown}", address=shown.lower())


class FakeReader:
    def __init__(self, latest: LatestMessage) -> None:
        self.latest = latest
        self.calls: list[tuple[str, str | None]] = []

    def latest_message(self, key: str, folder: str | None) -> LatestMessage:
        self.calls.append((key, folder))
        return self.latest


def latest_message(
    key: str,
    sender: Person | None = None,
    to: tuple[Person, ...] = (Person("Jane Doe", OWNER),),
    cc: tuple[Person, ...] = (),
    hidden: int = 0,
    subject: str = "Project Atlas budget",
) -> LatestMessage:
    return LatestMessage(
        key=key,
        subject=subject,
        sender=sender if sender is not None else Person("Omar Haddad", "omar@example.com"),
        to=to,
        cc=cc,
        hidden_recipients=hidden,
        body_top="Could you send the numbers?",
        sent="Today at 9:15",
        folder="Inbox",
    )


def sent_row(subject: str, time_label: str = "9:25 AM", who: str = "Omar Haddad") -> Node:
    return outlook_row(f"{who}, {subject},     {time_label},        Hello there")


@dataclass
class ComposeRig:
    service: ComposeService
    compose: FakeCompose
    outlook: FakeOutlookGateway
    guard: RecordingGuard
    reader: FakeReader
    clock: FakeClock
    scope_log: list[str]
    sends: SendGuard
    registry: RefRegistry

    def deadline(self, seconds: float = 45.0) -> Deadline:
        return Deadline.after(self.clock, seconds)

    def deliver_to_sent(self, subject: str, time_label: str = "9:27 AM") -> None:
        row = sent_row(subject, time_label=time_label, who="Bob Sample")
        rows = self.outlook.folder_rows_by_name.setdefault(SENT_FOLDER, [])
        rows.insert(0, row)
        if self.outlook.folder == SENT_FOLDER:
            self.outlook.rows.insert(0, row)


def compose_rig(
    compose: FakeCompose | None = None,
    guard: RecordingGuard | None = None,
    latest: LatestMessage | None = None,
    allowlist: str | None = None,
    sent_rows: list[Node] | None = None,
    owner: str | None = OWNER,
    scope: tuple[ScopeFactory, list[str]] | None = None,
) -> ComposeRig:
    clock = FakeClock(1000.0)
    waiter = fake_waiter(clock)
    inbox = outlook_row(INBOX_ROW, handle="inbox-row")
    outlook = FakeOutlookGateway(
        rows=[inbox],
        folders=[
            folder_row("Inbox; 1 unread message", handle="folder-inbox"),
            folder_row(SENT_FOLDER, handle="folder-sent"),
        ],
    )
    outlook.folder_rows_by_name = {"Inbox": [inbox], SENT_FOLDER: list(sent_rows or [])}
    registry = RefRegistry()
    rows = RowFinder(outlook, waiter, lambda: FAKE_TODAY, registry)
    fake = compose or FakeCompose()
    recording = guard or RecordingGuard()
    reader = FakeReader(latest or latest_message(inbox_key()))
    factory, log = scope or recording_factory()
    sends = SendGuard(clock)
    service = ComposeService(
        ComposeDeps(
            gateway=fake,
            rows=rows,
            scope=factory,
            reader=reader,
            guard=recording,
            recipients=RecipientGuard(parse_allowlist(allowlist), "48:notes"),
            sends=sends,
            waiter=waiter,
            now=lambda: FAKE_NOW,
            owner_email=lambda: owner,
        )
    )
    return ComposeRig(service, fake, outlook, recording, reader, clock, log, sends, registry)


def inbox_key() -> str:
    return parse_rows([outlook_row(INBOX_ROW)], FAKE_TODAY)[0].summary.key


class ExitFailureScope(RecordingScope):
    def __init__(self, options: ScopeOptions, log: list[str], failure: str) -> None:
        super().__init__(options, log)
        self.failure = failure

    def __exit__(self, *exc: object) -> None:
        super().__exit__(*exc)
        if len(exc) < 2 or exc[1] is None:
            raise CleanupError(self.failure)


def failing_exit_factory(failure: str) -> tuple[ScopeFactory, list[str]]:
    log: list[str] = []

    def factory(options: ScopeOptions) -> AbstractContextManager[UiScope]:
        return ExitFailureScope(options, log, failure)

    return factory, log
