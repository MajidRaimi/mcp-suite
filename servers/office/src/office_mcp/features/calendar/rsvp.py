from __future__ import annotations

from collections.abc import Callable

from office_mcp.features.calendar.cards import EventCards
from office_mcp.features.calendar.gateway import CalendarGateway
from office_mcp.features.calendar.locations import require_event_id
from office_mcp.features.calendar.models import MyResponse, RsvpOutcome, RsvpRequest, RsvpResponse
from office_mcp.features.calendar.peek import (
    ACCEPT_ID,
    DECLINE_ID,
    MESSAGE_PLACEHOLDER,
    MORE_ID,
    TENTATIVE_ITEM,
    by_dom_id,
    checkbox_state,
    email_checkbox,
    find_peek,
    message_area,
    my_response,
    respond_form,
    series_prompt,
)
from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.drafts import Channel, Draft, Review, TextGuard
from office_mcp.shared.errors import (
    ActionNotVerifiedError,
    ComposeBusyError,
    ElementNotFoundError,
    RecipientNotAllowedError,
    SendRefusedError,
    TargetError,
    TextEntryError,
)
from office_mcp.shared.menus import find_item
from office_mcp.shared.node import Node
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.waiting import Waiter

MAX_COMMENT = 2000
VERIFY_WAIT = 5.0
SERIES_WAIT = 1.0
NOTHING_SENT = "nothing was sent"
CLEAR_MESSAGE = "clear the message typed on the Teams event card"
MENU_ROLES = frozenset({"AXMenu", "AXList", "AXListBox"})
ALLOWLIST_REFUSAL = (
    "Not sent: OFFICE_MCP_SEND_ONLY_TO is active and the organizer's address cannot be checked "
    "from the Teams calendar, so calendar_respond only works with notify=false. Nothing was sent."
)
NO_FORM = "this event has no response buttons (it is not an invitation to you); nothing was sent"
UNCONFIRMED = (
    "the response was submitted but the event card did not confirm it; check the event in "
    "Teams before retrying"
)
SERIES_REFUSAL = (
    "Teams asked whether to answer one occurrence or the whole series; nothing was sent. "
    "Answer this one in Teams."
)
RESPONSE_IDS: dict[RsvpResponse, str] = {
    RsvpResponse.ACCEPT: ACCEPT_ID,
    RsvpResponse.DECLINE: DECLINE_ID,
}


def rsvp_request(
    event: str, response: str, notify: bool, comment: str | None, dry_run: bool
) -> RsvpRequest:
    event_id = require_event_id(event)
    try:
        choice = RsvpResponse(response)
    except ValueError as error:
        raise TargetError("response must be accept, tentative or decline") from error
    text = comment or ""
    if len(text) > MAX_COMMENT:
        raise TargetError(f"the comment is longer than {MAX_COMMENT} characters")
    if text.strip() and not notify:
        raise TargetError("a comment only reaches the organizer with notify=true")
    return RsvpRequest(event_id, choice, notify, text if text.strip() else "", dry_run)


class RsvpPolicy:
    def __init__(self, recipients: RecipientGuard, guard: TextGuard) -> None:
        self._recipients = recipients
        self._guard = guard

    def review(self, request: RsvpRequest) -> Review | None:
        if self._recipients.active and request.notify:
            raise RecipientNotAllowedError(ALLOWLIST_REFUSAL)
        if not request.comment:
            return None
        return self._guard.require_clean(Draft(channel=Channel.RSVP, body=request.comment))


def done_detail(target: MyResponse, notify: bool) -> str:
    if notify:
        return f"{target.value}; the organizer was emailed"
    return f"{target.value} without emailing the organizer"


class Responder:
    def __init__(
        self,
        gateway: CalendarGateway,
        waiter: Waiter,
        cards: EventCards,
        guard: TextGuard,
        reopen: Callable[[], Node | None],
    ) -> None:
        self._gateway = gateway
        self._waiter = waiter
        self._cards = cards
        self._guard = guard
        self._reopen = reopen
        self._committed = False

    def respond(
        self, cleanup: CleanupStack, peek: Node, request: RsvpRequest, review: Review | None
    ) -> RsvpOutcome:
        target = request.response.state
        if my_response(peek) is target:
            return RsvpOutcome(False, True, f"already {target.value}; {NOTHING_SENT}")
        if respond_form(peek) is None:
            raise TargetError(NO_FORM)
        peek = self._set_notify(peek, request.notify)
        if request.notify:
            self._fill_message(cleanup, peek, review)
        commit = self._commit_target(peek, request.response)
        self._gateway.press(commit)
        self._committed = True
        self._refuse_series_prompt()
        if not self._verify(target):
            return RsvpOutcome(True, False, UNCONFIRMED)
        return RsvpOutcome(True, True, done_detail(target, request.notify))

    def _set_notify(self, peek: Node, notify: bool) -> Node:
        box = email_checkbox(peek)
        if box is None:
            raise ElementNotFoundError(f"the event card has no Email organizer box; {NOTHING_SENT}")
        state = checkbox_state(box)
        if state is None:
            raise ActionNotVerifiedError(
                f"could not read the Email organizer box on the event card; {NOTHING_SENT}"
            )
        if state == notify:
            return peek
        self._gateway.press(box)
        return self._waiter.require(
            lambda: self._peek_with_notify(notify),
            lambda: ActionNotVerifiedError(
                f"the Email organizer box did not change; {NOTHING_SENT}"
            ),
        )

    def _peek_with_notify(self, notify: bool) -> Node | None:
        peek = find_peek(self._gateway.snapshot())
        box = email_checkbox(peek) if peek is not None else None
        if box is None or checkbox_state(box) != notify:
            return None
        return peek

    def _fill_message(self, cleanup: CleanupStack, peek: Node, review: Review | None) -> None:
        area = message_area(peek)
        if area is None:
            if review is None:
                return
            raise ElementNotFoundError(f"the event card has no message box; {NOTHING_SENT}")
        if self._gateway.read_text(area, MESSAGE_PLACEHOLDER).strip():
            raise ComposeBusyError(
                f"the message box on the Teams event card already has text; {NOTHING_SENT}"
            )
        if review is None:
            return
        text = review.draft.body
        self._gateway.write_text(area, text)
        cleanup.push(CLEAR_MESSAGE, lambda: self._clear(area, text))
        written = self._gateway.read_text(area, MESSAGE_PLACEHOLDER)
        self._guard.verify_written(review.draft, written, "")

    def _clear(self, area: Node, text: str) -> None:
        if self._committed:
            return
        if not self._gateway.clear_text(area, text, MESSAGE_PLACEHOLDER):
            raise TextEntryError("the message typed on the Teams event card could not be removed")

    def _commit_target(self, peek: Node, response: RsvpResponse) -> Node:
        if response is RsvpResponse.TENTATIVE:
            return self._tentative_item(peek)
        button = by_dom_id(peek, RESPONSE_IDS[response])
        if button is None:
            raise ElementNotFoundError(
                f"the event card has no {response.value} button; {NOTHING_SENT}"
            )
        return button

    def _tentative_item(self, peek: Node) -> Node:
        more = by_dom_id(peek, MORE_ID)
        if more is None:
            raise ElementNotFoundError(f"the event card has no More options button; {NOTHING_SENT}")
        self._gateway.press(more)
        return self._waiter.require(
            lambda: tentative_item(self._gateway.snapshot()),
            lambda: ElementNotFoundError(
                f"no Tentative option appeared under More options; {NOTHING_SENT}"
            ),
        )

    def _refuse_series_prompt(self) -> None:
        if self._waiter.until(lambda: series_prompt(self._gateway.snapshot()), SERIES_WAIT):
            self._gateway.close_popups()
            raise SendRefusedError(SERIES_REFUSAL)

    def _verify(self, target: MyResponse) -> bool:
        if self._waiter.until(lambda: self._shows(target), VERIFY_WAIT):
            return True
        peek = self._reopen()
        if peek is None:
            return target is MyResponse.DECLINED
        return my_response(peek) is target

    def _shows(self, target: MyResponse) -> bool:
        peek = self._cards.current()
        return peek is not None and my_response(peek) is target


def tentative_item(root: Node) -> Node | None:
    for menu in (node for node in root.walk() if node.role in MENU_ROLES):
        item = find_item(menu, (TENTATIVE_ITEM,))
        if item is not None:
            return item
    return None
