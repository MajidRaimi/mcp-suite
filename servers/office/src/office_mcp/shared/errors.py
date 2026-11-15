from __future__ import annotations

from collections.abc import Sequence

from office_mcp.shared.drafts import Severity, Violation, show_dashes


class OfficeError(Exception):
    pass


class SettingsError(OfficeError):
    pass


class PlatformError(OfficeError):
    pass


class AccessibilityError(OfficeError):
    pass


class AppNotRunningError(OfficeError):
    pass


class OutlookUnavailableError(OfficeError):
    pass


class AccessibilityModeError(OfficeError):
    pass


class ElementNotFoundError(OfficeError):
    pass


class AxActionError(OfficeError):
    pass


class MenuItemDisabledError(OfficeError):
    pass


class MenuStuckError(OfficeError):
    pass


class ActivationError(OfficeError):
    pass


class BusyError(OfficeError):
    pass


class SessionUnavailableError(OfficeError):
    pass


class UiTimeoutError(OfficeError):
    pass


class TargetError(OfficeError):
    pass


class TargetNotFoundError(TargetError):
    pass


class AmbiguousTargetError(TargetError):
    def __init__(self, message: str, candidates: Sequence[str]) -> None:
        super().__init__(message)
        self.candidates: tuple[str, ...] = tuple(candidates)

    def __str__(self) -> str:
        message = super().__str__()
        if not self.candidates:
            return message
        listed = "; ".join(self.candidates)
        return f"{message} Candidates: {listed}"


class ForbiddenActionError(OfficeError):
    pass


class TextEntryError(OfficeError):
    pass


class ActionNotVerifiedError(OfficeError):
    pass


class CleanupError(OfficeError):
    pass


class ComposeError(OfficeError):
    pass


class ComposeBusyError(ComposeError):
    pass


class SendRefusedError(OfficeError):
    pass


class DuplicateSendError(OfficeError):
    pass


class RecipientNotAllowedError(OfficeError):
    pass


class AttachmentError(OfficeError):
    pass


class PersonaError(OfficeError):
    pass


class PersonaMissingError(PersonaError):
    pass


class PersonaFormatError(PersonaError):
    pass


class PersonaViolationError(PersonaError):
    def __init__(self, violations: Sequence[Violation]) -> None:
        self.violations: tuple[Violation, ...] = tuple(violations)
        super().__init__(render_violations(self.violations))

    def __str__(self) -> str:
        return render_violations(self.violations)


class ToolDisabledError(OfficeError):
    pass


def error_text(error: OfficeError) -> str:
    notes = getattr(error, "__notes__", None) or []
    return "\n".join([str(error), *(str(note) for note in notes)])


def render_violations(violations: Sequence[Violation]) -> str:
    errors = [item for item in violations if item.severity is Severity.ERROR] or list(violations)
    noun = "rule" if len(errors) == 1 else "rules"
    header = (
        f"Not sent: the text breaks {len(errors)} persona {noun}. Fix them and call the tool again."
    )
    lines = [violation_line(number, item) for number, item in enumerate(errors, start=1)]
    return "\n".join([header, *lines])


def violation_line(number: int, violation: Violation) -> str:
    message = violation.message.rstrip()
    if violation.hint:
        ending = "" if message.endswith((".", "!", "?")) else "."
        message = f"{message}{ending} {violation.hint}"
    return show_dashes(f"{number}. {violation.rule}: {message}")
