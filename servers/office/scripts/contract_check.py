from __future__ import annotations

import enum
import importlib
import inspect
import sys
import typing
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from enum import StrEnum

SHARED = "office_mcp.shared"
OUTLOOK = "office_mcp.features.outlook"
TEAMS = "office_mcp.features.teams"
PENDING_SECTIONS = frozenset({"3.15", "3.16"})
SHARED_ONLY_FLAG = "--shared-only"


class Kind(StrEnum):
    SIGNATURE = "signature"
    EXISTS = "exists"
    PROPERTY = "property"
    BASE = "base"
    MEMBERS = "members"


class Status(StrEnum):
    OK = "ok"
    MISSING = "missing"
    WRONG = "wrong"
    PENDING = "pending F.2"


@dataclass(frozen=True)
class Expectation:
    section: str
    target: str
    kind: Kind
    expected: str = ""


@dataclass(frozen=True)
class Outcome:
    expectation: Expectation
    status: Status
    detail: str = ""

    def render(self) -> str:
        line = f"[{self.status.value}] {self.expectation.section} {self.expectation.target}"
        return f"{line}: {self.detail}" if self.detail else line


def sig(section: str, target: str, expected: str) -> Expectation:
    return Expectation(section, target, Kind.SIGNATURE, expected)


def exists(section: str, target: str) -> Expectation:
    return Expectation(section, target, Kind.EXISTS)


def prop(section: str, target: str) -> Expectation:
    return Expectation(section, target, Kind.PROPERTY)


def base(section: str, target: str, parent: str) -> Expectation:
    return Expectation(section, target, Kind.BASE, parent)


def members(section: str, target: str, expected: str) -> Expectation:
    return Expectation(section, target, Kind.MEMBERS, expected)


def errors_table() -> list[Expectation]:
    module = f"{SHARED}.errors"
    office = [
        "SettingsError",
        "PlatformError",
        "AccessibilityError",
        "AppNotRunningError",
        "OutlookUnavailableError",
        "AccessibilityModeError",
        "ElementNotFoundError",
        "AxActionError",
        "MenuItemDisabledError",
        "MenuStuckError",
        "ActivationError",
        "BusyError",
        "SessionUnavailableError",
        "UiTimeoutError",
        "TargetError",
        "ForbiddenActionError",
        "TextEntryError",
        "ActionNotVerifiedError",
        "CleanupError",
        "ComposeError",
        "SendRefusedError",
        "DuplicateSendError",
        "RecipientNotAllowedError",
        "AttachmentError",
        "PersonaError",
        "ToolDisabledError",
    ]
    nested = {
        "TargetNotFoundError": "TargetError",
        "AmbiguousTargetError": "TargetError",
        "ComposeBusyError": "ComposeError",
        "PersonaMissingError": "PersonaError",
        "PersonaFormatError": "PersonaError",
        "PersonaViolationError": "PersonaError",
    }
    return [
        base("3.1", f"{module}:OfficeError", "Exception"),
        *(base("3.1", f"{module}:{name}", "OfficeError") for name in office),
        *(base("3.1", f"{module}:{name}", parent) for name, parent in nested.items()),
        sig(
            "3.1",
            f"{module}:AmbiguousTargetError.__init__",
            "(self, message: str, candidates: Sequence[str]) -> None",
        ),
        sig(
            "3.1",
            f"{module}:PersonaViolationError.__init__",
            "(self, violations: Sequence[Violation]) -> None",
        ),
        sig("3.1", f"{module}:error_text", "(error: OfficeError) -> str"),
    ]


def drafts_table() -> list[Expectation]:
    module = f"{SHARED}.drafts"
    return [
        members("3.2", f"{module}:Channel", "EMAIL=email,TEAMS=teams,RSVP=rsvp"),
        members("3.2", f"{module}:Severity", "ERROR=error,WARNING=warning"),
        sig(
            "3.2",
            f"{module}:Violation",
            "(rule: str, message: str, severity: Severity, excerpt: str = '', hint: str = '')"
            " -> None",
        ),
        sig("3.2", f"{module}:Violation.to_record", "(self) -> dict[str, Any]"),
        sig(
            "3.2",
            f"{module}:Draft",
            "(channel: Channel, body: str, subject: str = '', context: str = '', "
            "context_author: str = '', sender_account: str = '', "
            "recipients: tuple[str, ...] = (), allow_language_mismatch: bool = False) -> None",
        ),
        sig("3.2", f"{module}:Fix", "(rule: str, detail: str) -> None"),
        sig(
            "3.2",
            f"{module}:Review",
            "(draft: Draft, violations: tuple[Violation, ...], fixes: tuple[Fix, ...], "
            "persona_version: str) -> None",
        ),
        prop("3.2", f"{module}:Review.ok"),
        prop("3.2", f"{module}:Review.warnings"),
        sig("3.2", f"{module}:Review.to_record", "(self) -> dict[str, Any]"),
        sig("3.2", f"{module}:TextGuard.review", "(self, draft: Draft) -> Review"),
        sig("3.2", f"{module}:TextGuard.require_clean", "(self, draft: Draft) -> Review"),
        sig(
            "3.2",
            f"{module}:TextGuard.verify_written",
            "(self, intended: Draft, body_read: str, subject_read: str) -> None",
        ),
        sig("3.2", f"{module}:NoTextGuard.review", "(self, draft: Draft) -> Review"),
        sig("3.2", f"{module}:NoTextGuard.require_clean", "(self, draft: Draft) -> Review"),
        sig(
            "3.2",
            f"{module}:NoTextGuard.verify_written",
            "(self, intended: Draft, body_read: str, subject_read: str) -> None",
        ),
        exists("3.2", f"{module}:NoTextGuard.persona_version"),
        sig("3.2", f"{module}:canonical_text", "(text: str) -> str"),
        sig("3.2", f"{module}:written_exactly", "(intended: str, actual: str) -> bool"),
    ]


def settings_table() -> list[Expectation]:
    module = f"{SHARED}.settings"
    return [
        exists("3.3", f"{module}:Settings.from_env"),
        sig("3.3", f"{module}:parse_groups", "(raw: str | None) -> frozenset[ToolGroup]"),
        sig("3.3", f"{module}:parse_bool", "(raw: str | None, name: str, default: bool) -> bool"),
        sig("3.3", f"{module}:parse_names", "(raw: str | None) -> frozenset[str]"),
        sig("3.3", f"{module}:parse_path", "(raw: str | None, default: Path) -> Path"),
        sig(
            "3.3",
            f"{module}:parse_seconds",
            "(raw: str | None, name: str, default: float, minimum: float, maximum: float) -> float",
        ),
        sig(
            "3.3",
            f"{module}:Settings",
            "(zone: Zone, host: str, port: int, transport: Transport, token: str | None, "
            "allowed_hosts: tuple[str, ...], owner_name: str | None, tool_policy: ToolPolicy, "
            "persona_file: Path, persona_explicit: bool, persona_required: bool, "
            "send_only_to: Allowlist | None, self_chat_thread: str, restore_frontmost: bool, "
            "action_timeout: float, idle_seconds: float, lock_file: Path, lock_timeout: float, "
            "stale_after: float, download_dir: Path | None, teams_max_members: int, "
            "owner_email: str | None) -> None",
        ),
    ]


def registry_table() -> list[Expectation]:
    module = f"{SHARED}.registry"
    return [
        members("3.4", f"{module}:ToolGroup", "READ=read,ORGANIZE=organize,SEND=send"),
        members(
            "3.4",
            f"{module}:Effect",
            "READS=reads,MARKS_READ=marks_read,ORGANIZES=organizes,SOFT_DELETES=soft_deletes,"
            "WRITES_LOCAL=writes_local,SENDS=sends",
        ),
        members("3.4", f"{module}:Stability", "STABLE=stable,EXPERIMENTAL=experimental"),
        sig(
            "3.4",
            f"{module}:ToolPolicy",
            "(groups: frozenset[ToolGroup], disabled: frozenset[str], "
            "experimental: frozenset[str], all_experimental: bool) -> None",
        ),
        sig("3.4", f"{module}:ToolPolicy.allows", "(self, spec: ToolSpec) -> bool"),
        sig(
            "3.4",
            f"{module}:ToolSpec",
            "(name: str, function: Callable[..., object], description: str, group: ToolGroup, "
            "effect: Effect, stability: Stability = Stability.STABLE) -> None",
        ),
        exists("3.4", f"{module}:EXPERIMENTAL_PREFIX"),
        sig("3.4", f"{module}:annotations_for", "(effect: Effect) -> ToolAnnotations"),
        sig(
            "3.4",
            f"{module}:enabled_specs",
            "(specs: Iterable[ToolSpec], policy: ToolPolicy) -> list[ToolSpec]",
        ),
        sig(
            "3.4",
            f"{module}:unknown_names",
            "(specs: Iterable[ToolSpec], policy: ToolPolicy) -> tuple[str, ...]",
        ),
    ]


def recipients_table() -> list[Expectation]:
    module = f"{SHARED}.recipients"
    return [
        exists("3.5", f"{module}:SELF_ENTRY"),
        sig(
            "3.5",
            f"{module}:Allowlist",
            "(emails: frozenset[str], domains: frozenset[str], teams_threads: frozenset[str], "
            "teams_self: bool) -> None",
        ),
        sig("3.5", f"{module}:Allowlist.entries", "(self) -> tuple[str, ...]"),
        sig("3.5", f"{module}:is_email", "(text: str) -> bool"),
        sig("3.5", f"{module}:normalize_email", "(text: str) -> str"),
        sig("3.5", f"{module}:parse_allowlist", "(raw: str | None) -> Allowlist | None"),
        sig(
            "3.5",
            f"{module}:RecipientGuard.__init__",
            "(self, allowlist: Allowlist | None, self_thread: str) -> None",
        ),
        prop("3.5", f"{module}:RecipientGuard.active"),
        sig(
            "3.5",
            f"{module}:RecipientGuard.require_emails",
            "(self, addresses: Sequence[str]) -> None",
        ),
        sig("3.5", f"{module}:RecipientGuard.require_thread", "(self, thread_id: str) -> None"),
        sig("3.5", f"{module}:RecipientGuard.describe", "(self) -> list[str] | None"),
    ]


def events_table() -> list[Expectation]:
    module = f"{SHARED}.events"
    menus = f"{SHARED}.menus"
    return [
        members(
            "3.6",
            f"{module}:Key",
            "A=0,S=1,D=2,F=3,G=5,C=8,V=9,Q=12,E=14,R=15,T=17,O=31,J=38,N=45,M=46,RETURN=36,"
            "TAB=48,SPACE=49,DELETE=51,ESCAPE=53,LEFT=123,RIGHT=124,DOWN=125,UP=126",
        ),
        exists("3.6", f"{module}:Modifier.COMMAND"),
        exists("3.6", f"{module}:Modifier.SHIFT"),
        exists("3.6", f"{module}:Modifier.OPTION"),
        exists("3.6", f"{module}:Modifier.CONTROL"),
        sig(
            "3.6",
            f"{module}:Shortcut",
            "(key: Key, modifiers: Modifier = Modifier.NONE) -> None",
        ),
        sig("3.6", f"{module}:InputEvents.press", "(self, pid: int, shortcut: Shortcut) -> None"),
        sig(
            "3.6",
            f"{module}:InputEvents.click",
            "(self, pid: int, point: tuple[float, float]) -> None",
        ),
        sig(
            "3.6",
            f"{module}:GuardedEvents.__init__",
            "(self, inner: InputEvents, forbidden: frozenset[Shortcut]) -> None",
        ),
        sig("3.6", f"{module}:GuardedEvents.press", "(self, pid: int, shortcut: Shortcut) -> None"),
        sig(
            "3.6",
            f"{module}:GuardedEvents.click",
            "(self, pid: int, point: tuple[float, float]) -> None",
        ),
        sig("3.6", f"{module}:decode_menu_modifiers", "(value: int) -> Modifier"),
        exists("3.6", f"{module}:OUTLOOK_FORBIDDEN_SHORTCUTS"),
        exists("3.6", f"{module}:TEAMS_FORBIDDEN_SHORTCUTS"),
        sig(
            "3.6",
            f"{menus}:MenuEntry",
            "(path: tuple[str, ...], enabled: bool, checked: bool, shortcut: Shortcut | None, "
            "handle: object) -> None",
        ),
        sig("3.6", f"{menus}:MenuBar.__init__", "(self, app: AxApp) -> None"),
        sig("3.6", f"{menus}:MenuBar.entries", "(self, top: str) -> list[MenuEntry]"),
        sig(
            "3.6",
            f"{menus}:MenuBar.entry",
            "(self, path: tuple[str, ...], shortcut: Shortcut | None = None) -> MenuEntry",
        ),
        sig(
            "3.6",
            f"{menus}:MenuBar.press",
            "(self, path: tuple[str, ...], shortcut: Shortcut | None = None) -> None",
        ),
        sig(
            "3.6",
            f"{menus}:PopupMenus.__init__",
            "(self, app: AxApp, events: InputEvents, waiter: Waiter) -> None",
        ),
        sig(
            "3.6",
            f"{menus}:PopupMenus.open",
            "(self, element: object, action: str = 'AXShowMenu') -> Node",
        ),
        sig(
            "3.6",
            f"{menus}:PopupMenus.choose",
            "(self, menu: Node, path: tuple[str, ...]) -> None",
        ),
        sig(
            "3.6",
            f"{menus}:PopupMenus.checked",
            "(self, menu: Node, path: tuple[str, ...]) -> bool | None",
        ),
        sig("3.6", f"{menus}:PopupMenus.close", "(self, menu: Node | None = None) -> None"),
        sig("3.6", f"{menus}:find_item", "(menu: Node, path: tuple[str, ...]) -> Node | None"),
        sig("3.6", f"{menus}:is_forbidden", "(path: Sequence[str]) -> bool"),
        exists("3.6", f"{menus}:FORBIDDEN_MENU_WORDS"),
    ]


def ax_table() -> list[Expectation]:
    ax = f"{SHARED}.ax"
    app = f"{ax}:AxApp"
    return [
        sig("3.7", f"{ax}:AxBackend.is_trusted", "(self) -> bool"),
        sig("3.7", f"{ax}:AxBackend.application", "(self, pid: int) -> object"),
        sig(
            "3.7",
            f"{ax}:AxBackend.copy",
            "(self, element: object, name: str) -> tuple[int, object]",
        ),
        sig(
            "3.7",
            f"{ax}:AxBackend.copy_many",
            "(self, element: object, names: Sequence[str]) -> tuple[int, Sequence[object] | None]",
        ),
        sig(
            "3.7", f"{ax}:AxBackend.set", "(self, element: object, name: str, value: object) -> int"
        ),
        sig("3.7", f"{ax}:AxBackend.perform", "(self, element: object, action: str) -> int"),
        sig("3.7", f"{ax}:AxBackend.action_names", "(self, element: object) -> tuple[str, ...]"),
        sig("3.7", f"{ax}:AxBackend.same", "(self, first: object, second: object) -> bool"),
        sig("3.7", f"{ax}:AxBackend.range_value", "(self, location: int, length: int) -> object"),
        sig("3.7", f"{ax}:ProcessDirectory.pid", "(self, bundle_id: str) -> int | None"),
        sig("3.7", f"{ax}:ProcessDirectory.launch", "(self, app_name: str) -> None"),
        members("3.7", f"{ax}:AccessMode", "NATIVE=native,WEB=web,WEB_ENHANCED=web_enhanced"),
        exists("3.7", f"{ax}:MENU_OPEN_CODES"),
        sig(
            "3.7",
            f"{app}.__init__",
            "(self, backend: AxBackend, processes: ProcessDirectory, bundle_id: str, "
            "app_name: str, waiter: Waiter, mode: AccessMode = AccessMode.NATIVE) -> None",
        ),
        sig("3.7", f"{app}.trusted", "(self) -> bool"),
        sig("3.7", f"{app}.require_trust", "(self) -> None"),
        sig("3.7", f"{app}.pid", "(self) -> int | None"),
        sig("3.7", f"{app}.require_pid", "(self) -> int"),
        sig("3.7", f"{app}.running", "(self) -> bool"),
        sig("3.7", f"{app}.application", "(self) -> object"),
        sig("3.7", f"{app}.windows", "(self) -> list[object]"),
        sig("3.7", f"{app}.attribute", "(self, element: object, name: str) -> object | None"),
        sig("3.7", f"{app}.text", "(self, element: object, name: str) -> str"),
        sig(
            "3.7",
            f"{app}.set_attribute",
            "(self, element: object, name: str, value: object, "
            "tolerate: frozenset[int] = frozenset()) -> None",
        ),
        sig(
            "3.7",
            f"{app}.perform",
            "(self, element: object, action: str, label: str = '', "
            "tolerate: frozenset[int] = frozenset()) -> int",
        ),
        sig("3.7", f"{app}.actions", "(self, element: object) -> tuple[str, ...]"),
        sig("3.7", f"{app}.custom_action", "(self, element: object, label: str) -> str | None"),
        sig(
            "3.7",
            f"{app}.snapshot",
            "(self, element: object, max_depth: int = 90, "
            "prune: Callable[[Node], bool] | None = None) -> Node",
        ),
        sig("3.7", f"{app}.same", "(self, first: object, second: object) -> bool"),
        sig("3.7", f"{app}.set_enhanced", "(self, value: bool) -> bool | None"),
        exists("3.7", f"{app}.window_or_launch"),
        sig(
            "3.7",
            f"{SHARED}.snapshot:Snapshotter.__init__",
            "(self, app: AxApp, max_depth: int = 90, prune: Callable[[Node], bool] | None = None, "
            "budget: int = 40000) -> None",
        ),
        sig(
            "3.7",
            f"{SHARED}.tree:find",
            "(app: AxApp, root: object, matches: Callable[[object], bool], max_depth: int = 30, "
            "prune: Callable[[object], bool] | None = None) -> object | None",
        ),
        sig(
            "3.7",
            f"{SHARED}.tree:find_all",
            "(app: AxApp, root: object, matches: Callable[[object], bool], max_depth: int = 30, "
            "prune: Callable[[object], bool] | None = None, limit: int = 500) -> list[object]",
        ),
        sig(
            "3.7",
            f"{SHARED}.tree:by_role_and_label",
            "(app: AxApp, role: str, label: str) -> Callable[[object], bool]",
        ),
        sig(
            "3.7",
            f"{SHARED}.windows:WindowProbe",
            "(handle: object, title: str, subrole: str, main: bool, focused: bool, "
            "landmarks: frozenset[str]) -> None",
        ),
        sig(
            "3.7",
            f"{SHARED}.windows:WindowLocator.__init__",
            "(self, app: AxApp, probe_depth: int = 5, "
            "prune_roles: frozenset[str] = PRUNED_ROLES) -> None",
        ),
        sig("3.7", f"{SHARED}.windows:WindowLocator.probes", "(self) -> list[WindowProbe]"),
        sig(
            "3.7",
            f"{SHARED}.windows:WindowLocator.first",
            "(self, classify: Callable[[WindowProbe], K], kind: K) -> WindowProbe | None",
        ),
        sig(
            "3.7",
            f"{SHARED}.windows:WindowLocator.require",
            "(self, classify: Callable[[WindowProbe], K], kind: K, error: str) -> WindowProbe",
        ),
        sig(
            "3.7",
            f"{SHARED}.windows:WindowLocator.wait_new",
            "(self, before: Sequence[WindowProbe], classify: Callable[[WindowProbe], K], "
            "kind: K) -> WindowProbe",
        ),
        sig(
            "3.7",
            f"{SHARED}.node:Node",
            "(role: str, title: str = '', value: str = '', description: str = '', "
            "dom_id: str = '', classes: tuple[str, ...] = (), children: tuple[Node, ...] = (), "
            "handle: object | None = None, subrole: str = '', help: str = '', "
            "enabled: bool | None = None, level: int | None = None, "
            "selected: bool | None = None) -> None",
        ),
    ]


def waiting_table() -> list[Expectation]:
    waiting = f"{SHARED}.waiting:Waiter"
    runner = f"{SHARED}.runner"
    cleanup = f"{SHARED}.cleanup:CleanupStack"
    return [
        sig(
            "3.8",
            waiting,
            "(clock: Callable[[], float], sleep: Callable[[float], None], timeout: float = 12.0, "
            "poll: float = 0.4) -> None",
        ),
        sig(
            "3.8",
            f"{waiting}.until",
            "(self, ready: Callable[[], T | None], timeout: float | None = None) -> T | None",
        ),
        sig(
            "3.8",
            f"{waiting}.require",
            "(self, ready: Callable[[], T | None], failure: Callable[[], OfficeError], "
            "timeout: float | None = None) -> T",
        ),
        sig(
            "3.8",
            f"{waiting}.settle",
            "(self, sample: Callable[[], int], minimum: int = 0, stable_rounds: int = 2, "
            "timeout: float | None = None) -> int",
        ),
        sig("3.8", f"{waiting}.with_timeout", "(self, timeout: float) -> Waiter"),
        sig("3.8", f"{cleanup}.push", "(self, label: str, undo: Callable[[], None]) -> None"),
        sig("3.8", f"{cleanup}.__enter__", "(self) -> CleanupStack"),
        sig(
            "3.8",
            f"{cleanup}.__exit__",
            "(self, kind: type[BaseException] | None, error: BaseException | None, tb: object)"
            " -> None",
        ),
        prop("3.8", f"{cleanup}.report"),
        sig(
            "3.8",
            f"{runner}:Deadline",
            "(clock: Callable[[], float], expires_at: float) -> None",
        ),
        sig(
            "3.8",
            f"{runner}:Deadline.after",
            "(clock: Callable[[], float], seconds: float) -> Deadline",
        ),
        sig("3.8", f"{runner}:Deadline.remaining", "(self) -> float"),
        sig("3.8", f"{runner}:Deadline.expired", "(self) -> bool"),
        sig("3.8", f"{runner}:UiLock.acquire", "(self, timeout: float) -> bool"),
        sig("3.8", f"{runner}:UiLock.release", "(self) -> None"),
        sig(
            "3.8",
            f"{runner}:ProcessUiLock.__init__",
            "(self, path: Path, clock: Callable[[], float], sleep: Callable[[float], None])"
            " -> None",
        ),
        sig("3.8", f"{runner}:ProcessUiLock.acquire", "(self, timeout: float) -> bool"),
        sig("3.8", f"{runner}:ProcessUiLock.release", "(self) -> None"),
        exists("3.8", f"{runner}:READ_BUDGET"),
        exists("3.8", f"{runner}:WRITE_BUDGET"),
        exists("3.8", f"{runner}:SEND_BUDGET"),
        exists("3.8", f"{runner}:BATCH_BUDGET"),
        sig(
            "3.8",
            f"{runner}:Runner.__init__",
            "(self, lock: UiLock, clock: Callable[[], float], lock_timeout: float, "
            "stale_after: float) -> None",
        ),
        sig(
            "3.8",
            f"{runner}:Runner.read",
            "(self, work: Callable[[Deadline], T], budget: float = 120.0) -> T",
        ),
        sig(
            "3.8",
            f"{runner}:Runner.write",
            "(self, work: Callable[[Deadline], T], budget: float = 30.0) -> T",
        ),
        sig("3.8", f"{runner}:Lazy.__init__", "(self, build: Callable[[], T]) -> None"),
        sig("3.8", f"{runner}:Lazy.get", "(self) -> T"),
    ]


def focus_table() -> list[Expectation]:
    module = f"{SHARED}.focus"
    session = f"{module}:ForegroundSession"
    workspace = f"{module}:Workspace"
    return [
        members(
            "3.9",
            f"{module}:FocusOutcome",
            "NOT_NEEDED=not_needed,RESTORED=restored,OWNER_MOVED=owner_moved,KEPT=kept,"
            "FAILED=failed",
        ),
        sig("3.9", f"{workspace}.frontmost_pid", "(self) -> int | None"),
        sig("3.9", f"{workspace}.activate", "(self, pid: int) -> None"),
        sig("3.9", f"{workspace}.is_running", "(self, pid: int) -> bool"),
        sig("3.9", f"{workspace}.app_name", "(self, pid: int) -> str"),
        sig("3.9", f"{workspace}.idle_seconds", "(self) -> float"),
        sig("3.9", f"{workspace}.screen_locked", "(self) -> bool"),
        sig("3.9", f"{workspace}.secure_input", "(self) -> bool"),
        sig(
            "3.9",
            f"{module}:FocusPolicy",
            "(restore: bool, idle_seconds: float, idle_wait: float = 10.0) -> None",
        ),
        sig(
            "3.9",
            f"{session}.__init__",
            "(self, workspace: Workspace, app: AxApp, waiter: Waiter, policy: FocusPolicy) -> None",
        ),
        sig("3.9", f"{session}.__enter__", "(self) -> ForegroundSession"),
        sig("3.9", f"{session}.__exit__", "(self, *exc: object) -> None"),
        sig("3.9", f"{session}.ensure_front", "(self) -> None"),
        prop("3.9", f"{session}.outcome"),
    ]


def text_input_table() -> list[Expectation]:
    pasteboard = f"{SHARED}.pasteboard"
    writer = f"{SHARED}.text_input:TextWriter"
    return [
        sig("3.10", f"{pasteboard}:Pasteboard.snapshot", "(self) -> PasteboardSnapshot"),
        sig(
            "3.10",
            f"{pasteboard}:Pasteboard.restore",
            "(self, snapshot: PasteboardSnapshot) -> None",
        ),
        sig("3.10", f"{pasteboard}:Pasteboard.write_transient_text", "(self, text: str) -> None"),
        sig(
            "3.10",
            f"{pasteboard}:PasteboardSnapshot",
            "(items: tuple[Mapping[str, bytes], ...]) -> None",
        ),
        sig("3.10", f"{pasteboard}:preserved", "(pasteboard: Pasteboard) -> Iterator[None]"),
        sig(
            "3.10",
            f"{writer}.__init__",
            "(self, app: AxApp, events: InputEvents, pasteboard: Pasteboard, waiter: Waiter)"
            " -> None",
        ),
        sig("3.10", f"{writer}.read", "(self, element: object, placeholder: str = '') -> str"),
        sig("3.10", f"{writer}.set_value", "(self, element: object, text: str) -> None"),
        sig(
            "3.10",
            f"{writer}.write",
            "(self, element: object, text: str, at_start: bool = False, "
            "placeholder: str = '') -> None",
        ),
        sig(
            "3.10",
            f"{writer}.clear_exact",
            "(self, element: object, inserted: str, placeholder: str = '') -> bool",
        ),
    ]


def results_table() -> list[Expectation]:
    module = f"{SHARED}.results"
    return [
        sig(
            "3.11",
            f"{module}:ItemResult",
            "(ref: str, ok: bool, label: str, location: str, verified: bool, detail: str = '')"
            " -> None",
        ),
        sig("3.11", f"{module}:ItemResult.to_record", "(self) -> dict[str, Any]"),
        sig(
            "3.11",
            f"{module}:ActionResult",
            "(action: str, items: tuple[ItemResult, ...], skipped: tuple[str, ...], "
            "focus: FocusOutcome, notes: tuple[str, ...] = (), "
            "extra: Mapping[str, str] = <factory>) -> None",
        ),
        sig("3.11", f"{module}:ActionResult.to_record", "(self) -> dict[str, Any]"),
        members(
            "3.11",
            f"{module}:SendStatus",
            "SENT=sent,SUBMITTED=submitted,UNCONFIRMED=unconfirmed,DRY_RUN=dry_run",
        ),
        sig(
            "3.11",
            f"{module}:SendResult",
            "(status: SendStatus, channel: Channel, targets: tuple[str, ...], subject: str, "
            "message_ref: str, verified: bool, review: Review, focus: FocusOutcome, "
            "notes: tuple[str, ...] = (), extra: Mapping[str, Any] = <factory>) -> None",
        ),
        sig("3.11", f"{module}:SendResult.to_record", "(self) -> dict[str, Any]"),
    ]


def scope_owner_runtime_table() -> list[Expectation]:
    scope = f"{SHARED}.scope"
    phrases = f"{SHARED}.phrases"
    return [
        sig(
            "3.12", f"{scope}:ScopeOptions", "(foreground: bool, restore_view: bool = True) -> None"
        ),
        sig("3.12", f"{scope}:UiScope.ensure_front", "(self) -> None"),
        prop("3.12", f"{scope}:UiScope.focus"),
        exists("3.12", f"{scope}:ScopeFactory"),
        sig(
            "3.13",
            f"{SHARED}.owner:Owner",
            "(name: str | None = None, email: str | None = None) -> None",
        ),
        prop("3.13", f"{SHARED}.owner:Owner.label"),
        prop("3.13", f"{SHARED}.owner:Owner.possessive"),
        *(
            exists("3.13", f"{phrases}:{name}")
            for name in (
                "OUTLOOK_UNAVAILABLE",
                "FOREGROUND_NOTE",
                "INJECTION_RULE",
                "NO_RETRY_NOTE",
                "KEYS_ONLY_NOTE",
                "TEAMS_TARGET_NOTE",
            )
        ),
        sig(
            "3.14",
            f"{SHARED}.runtime:Runtime",
            "(settings: Settings, waiter: Waiter, today: Callable[[], date], "
            "now: Callable[[], datetime], backend: AxBackend, processes: ProcessDirectory, "
            "workspace: Workspace, outlook_events: InputEvents, teams_events: InputEvents, "
            "pasteboard: Pasteboard, recipients: RecipientGuard) -> None",
        ),
        sig("3.14", f"{SHARED}.runtime:build_runtime", "(settings: Settings) -> Runtime"),
    ]


def outlook_table() -> list[Expectation]:
    window = f"{OUTLOOK}.window"
    rows = f"{OUTLOOK}.rows"
    return [
        members("3.15", f"{window}:OutlookWindow", "MAIN=main,COMPOSE=compose,OTHER=other"),
        members(
            "3.15",
            f"{window}:OutlookMode",
            "MAIL=mail,SEARCH=search,CALENDAR=calendar,OTHER=other",
        ),
        members("3.15", f"{window}:OutlookModule", "MAIL=Mail,CALENDAR=Calendar"),
        members("3.15", f"{window}:FocusedTab", "FOCUSED=focused,OTHER=other"),
        sig("3.15", f"{window}:classify_window", "(probe: WindowProbe) -> OutlookWindow"),
        sig("3.15", f"{window}:detect_mode", "(probe: WindowProbe) -> OutlookMode"),
        sig("3.15", f"{window}:account_from_title", "(title: str) -> str | None"),
        sig(
            "3.15",
            f"{OUTLOOK}.models:Folder",
            "(name: str, unread: int, total: int | None = None, unseen: int | None = None, "
            "level: int = 0, path: str = '', favorite: bool = False) -> None",
        ),
        sig(
            "3.15",
            f"{OUTLOOK}.models:Summary",
            "(key: str, index: int, group: str, flags: tuple[str, ...], unread: bool, "
            "unread_count: int, message_count: int, title: str, meeting: str, time_label: str, "
            "date: str | None, preview: str, thread: tuple[str, ...] = (), folder: str = '', "
            "view: str | None = None, flagged: bool = False, replied: bool = False, "
            "forwarded: bool = False, has_attachments: bool = False, pinned: bool = False)"
            " -> None",
        ),
        sig("3.15", f"{OUTLOOK}.models:Attachment", "(name: str, size: str) -> None"),
        sig(
            "3.15",
            f"{OUTLOOK}.models:ThreadMessage",
            "(sender: str, sent: str, to: str, cc: str, subject: str, body: str) -> None",
        ),
        sig(
            "3.15",
            f"{OUTLOOK}.listing:parse_folder",
            "(description: str, level: int = 0) -> Folder | None",
        ),
        sig("3.15", f"{OUTLOOK}.listing:parse_folders", "(nodes: list[Node]) -> list[Folder]"),
        sig(
            "3.15",
            f"{OUTLOOK}.thread:split_thread",
            "(body: str, limit: int = 20) -> tuple[ThreadMessage, ...]",
        ),
        sig("3.15", f"{OUTLOOK}.thread:top_text", "(body: str) -> str"),
        sig(
            "3.15",
            f"{OUTLOOK}.conversation:ConversationReader.latest_message",
            "(self, key: str, folder: str | None) -> LatestMessage",
        ),
        members(
            "3.15",
            f"{rows}:SearchScope",
            "CURRENT_FOLDER=Current Folder,SUBFOLDERS=Subfolders,MAILBOX=Current Mailbox,"
            "ALL=All Mailboxes",
        ),
        sig("3.15", f"{rows}:is_key", "(ref: str) -> bool"),
        sig("3.15", f"{rows}:RefRegistry.__init__", "(self, capacity: int = 500) -> None"),
        sig(
            "3.15",
            f"{rows}:RowFinder.__init__",
            "(self, gateway: OutlookGateway, waiter: Waiter, today: Callable[[], date], "
            "registry: RefRegistry, front: FrontHook | None = None) -> None",
        ),
        sig(
            "3.15",
            f"{rows}:RowFinder.resolve_for_write",
            "(self, key: str, rows: list[Row]) -> Row",
        ),
        sig("3.15", f"{rows}:RowFinder.refresh", "(self, key: str, locator: Locator) -> Row"),
        sig(
            "3.15",
            f"{OUTLOOK}.view:capture_view",
            "(gateway: OutlookGateway) -> ViewState",
        ),
        sig(
            "3.15",
            f"{OUTLOOK}.scope:OutlookScope.__init__",
            "(self, driver: OutlookDriver, options: ScopeOptions) -> None",
        ),
        sig(
            "3.15",
            f"{OUTLOOK}.driver:build_outlook_driver",
            "(runtime: Runtime, registry: RefRegistry) -> OutlookDriver",
        ),
    ]


def teams_table() -> list[Expectation]:
    return [
        members(
            "3.16",
            f"{TEAMS}.rail:TeamsApp",
            "ACTIVITY=14d6962d-6eeb-4f48-8890-de55454bb136,"
            "CHAT=86fcd49b-61a2-4701-b771-54728cd291fb,"
            "TEAMS=2a84919f-59d8-4441-a975-2a8c2643b741,"
            "CALENDAR=ef56c0de-36fc-4ef8-b417-3d82ba9d073c",
        ),
        sig("3.16", f"{TEAMS}.rail:selected_app", "(root: Node) -> TeamsApp | None"),
        sig("3.16", f"{TEAMS}.header:open_thread_id", "(root: Node) -> str | None"),
        sig("3.16", f"{TEAMS}.header:participant_count", "(root: Node) -> int | None"),
        members(
            "3.16",
            f"{TEAMS}.header:ThreadKind",
            "CHAT=chat,GROUP=group,MEETING=meeting,CHANNEL=channel,NOTES=notes",
        ),
        sig("3.16", f"{TEAMS}.header:thread_kind", "(thread_id: str) -> ThreadKind"),
        sig("3.16", f"{TEAMS}.gateway:TeamsGateway.show_menu", "(self, node: Node) -> Node"),
        sig("3.16", f"{TEAMS}.gateway:TeamsGateway.close_popups", "(self) -> int"),
        sig(
            "3.16",
            f"{TEAMS}.scope:TeamsScope.__init__",
            "(self, driver: TeamsDriver, options: ScopeOptions) -> None",
        ),
        sig("3.16", f"{TEAMS}.scope:TeamsScope.show_app", "(self, app: TeamsApp) -> Node"),
        sig(
            "3.16",
            f"{TEAMS}.chats.opener:resolve_for_write",
            "(query: str, rows: list[ChatRow]) -> ChatRow",
        ),
        sig(
            "3.16",
            f"{TEAMS}.chats.opener:ChatOpener.open",
            "(self, query: str, for_write: bool) -> OpenedChat",
        ),
        sig(
            "3.16",
            f"{TEAMS}.messages.models:Quote",
            "(author: str, sent: str, text: str) -> None",
        ),
        sig(
            "3.16",
            f"{TEAMS}.messages.models:Message",
            "(id: str, author: str, mine: bool, sent_at: str, text: str, "
            "quote: Quote | None = None, edited: bool = False) -> None",
        ),
        sig("3.16", f"{TEAMS}.driver:build_teams_driver", "(runtime: Runtime) -> TeamsDriver"),
    ]


def table() -> list[Expectation]:
    return [
        *errors_table(),
        *drafts_table(),
        *settings_table(),
        *registry_table(),
        *recipients_table(),
        *events_table(),
        *ax_table(),
        *waiting_table(),
        *focus_table(),
        *text_input_table(),
        *results_table(),
        *scope_owner_runtime_table(),
        *outlook_table(),
        *teams_table(),
    ]


def resolve(target: str) -> object:
    module_name, _, qualname = target.partition(":")
    current: object = importlib.import_module(module_name)
    for part in qualname.split("."):
        current = inspect.getattr_static(current, part)
    return current


def unwrap(value: object) -> object:
    if isinstance(value, classmethod | staticmethod):
        return value.__func__
    return value


def annotation_text(annotation: object) -> str:
    return annotation if isinstance(annotation, str) else inspect.formatannotation(annotation)


def default_text(value: object) -> str:
    if isinstance(value, enum.Enum):
        return f"{type(value).__name__}.{value.name}"
    return repr(value)


def parameter_text(parameter: inspect.Parameter) -> str:
    prefix = {
        inspect.Parameter.VAR_POSITIONAL: "*",
        inspect.Parameter.VAR_KEYWORD: "**",
    }.get(parameter.kind, "")
    text = f"{prefix}{parameter.name}"
    if parameter.annotation is not inspect.Parameter.empty:
        text += f": {annotation_text(parameter.annotation)}"
    if parameter.default is not inspect.Parameter.empty:
        separator = " = " if parameter.annotation is not inspect.Parameter.empty else "="
        text += f"{separator}{default_text(parameter.default)}"
    return text


def signature_text(signature: inspect.Signature, skip_first: bool) -> str:
    parameters = list(signature.parameters.values())[1 if skip_first else 0 :]
    rendered: list[str] = []
    star_written = False
    for parameter in parameters:
        if parameter.kind is inspect.Parameter.VAR_POSITIONAL:
            star_written = True
        if parameter.kind is inspect.Parameter.KEYWORD_ONLY and not star_written:
            rendered.append("*")
            star_written = True
        rendered.append(parameter_text(parameter))
    text = f"({', '.join(rendered)})"
    if signature.return_annotation is not inspect.Signature.empty:
        text += f" -> {annotation_text(signature.return_annotation)}"
    return text


def candidate_texts(value: object) -> Iterator[str]:
    function = unwrap(value)
    skip_first = isinstance(value, classmethod)
    if inspect.isclass(function):
        yield signature_text(inspect.signature(function), skip_first=False)
        return
    if not callable(function):
        return
    for variant in (*typing.get_overloads(function), function):
        yield signature_text(inspect.signature(variant), skip_first)


def rendered_defaults(text: str, module_name: str) -> str:
    module = sys.modules.get(module_name)
    if module is None:
        return text
    for name, item in vars(module).items():
        if isinstance(item, frozenset) and item and repr(item) in text and name.isupper():
            text = text.replace(repr(item), name)
    return text


def check_signature(expectation: Expectation, value: object) -> Outcome:
    module_name = expectation.target.partition(":")[0]
    texts = [rendered_defaults(text, module_name) for text in candidate_texts(value)]
    if expectation.expected in texts:
        return Outcome(expectation, Status.OK)
    found = texts[-1] if texts else "not callable"
    return Outcome(expectation, Status.WRONG, f"expected {expectation.expected} found {found}")


def check_members(expectation: Expectation, value: object) -> Outcome:
    if not (inspect.isclass(value) and issubclass(value, enum.Enum)):
        return Outcome(expectation, Status.WRONG, "not an enum")
    found = ",".join(f"{item.name}={item.value}" for item in value)
    wanted = dict(pair.split("=", 1) for pair in expectation.expected.split(","))
    actual = {item.name: str(item.value) for item in value}
    if wanted == actual:
        return Outcome(expectation, Status.OK)
    return Outcome(expectation, Status.WRONG, f"expected {expectation.expected} found {found}")


def check_base(expectation: Expectation, value: object) -> Outcome:
    parents = [cls.__name__ for cls in getattr(value, "__mro__", ())[1:]]
    if expectation.expected in parents[:1]:
        return Outcome(expectation, Status.OK)
    return Outcome(
        expectation, Status.WRONG, f"expected base {expectation.expected}, bases {parents[:1]}"
    )


CHECKS: dict[Kind, Callable[[Expectation, object], Outcome]] = {
    Kind.SIGNATURE: check_signature,
    Kind.MEMBERS: check_members,
    Kind.BASE: check_base,
    Kind.EXISTS: lambda expectation, value: Outcome(expectation, Status.OK),
    Kind.PROPERTY: lambda expectation, value: Outcome(
        expectation, Status.OK if isinstance(value, property) else Status.WRONG, "not a property"
    ),
}


def evaluate(expectation: Expectation) -> Outcome:
    try:
        value = resolve(expectation.target)
    except (ImportError, AttributeError) as error:
        status = Status.PENDING if expectation.section in PENDING_SECTIONS else Status.MISSING
        return Outcome(expectation, status, str(error))
    outcome = CHECKS[expectation.kind](expectation, value)
    if outcome.status is Status.OK:
        return Outcome(expectation, Status.OK)
    return outcome


def main(arguments: list[str]) -> int:
    shared_only = SHARED_ONLY_FLAG in arguments
    expectations = [
        item for item in table() if not (shared_only and item.section in PENDING_SECTIONS)
    ]
    outcomes = [evaluate(item) for item in expectations]
    for outcome in outcomes:
        if outcome.status is not Status.OK:
            print(outcome.render())
    failures = [item for item in outcomes if item.status in (Status.WRONG, Status.MISSING)]
    pending = [item for item in outcomes if item.status is Status.PENDING]
    ok = len(outcomes) - len(failures) - len(pending)
    print(f"contract: {ok} ok, {len(pending)} pending F.2, {len(failures)} failing")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
