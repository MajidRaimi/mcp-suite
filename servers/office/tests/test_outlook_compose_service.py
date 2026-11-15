from __future__ import annotations

import pytest
from fakes.guard import RecordingGuard
from fakes.p3_compose import (
    OWNER,
    SENT_FOLDER,
    FakeCompose,
    compose_rig,
    failing_exit_factory,
    inbox_key,
    latest_message,
    sent_row,
)

from office_mcp.features.outlook.compose.gateway import (
    AlertSheet,
    RecipientField,
    RecipientToken,
)
from office_mcp.features.outlook.compose.service import ForwardMail, NewMail, ReplyMail
from office_mcp.features.outlook.models import Person
from office_mcp.shared.drafts import Channel, Draft, Fix, Review, Severity, Violation
from office_mcp.shared.errors import (
    AxActionError,
    ComposeBusyError,
    ComposeError,
    DuplicateSendError,
    PersonaViolationError,
    RecipientNotAllowedError,
    SendRefusedError,
    TargetError,
    TextEntryError,
    UiTimeoutError,
    error_text,
)
from office_mcp.shared.node import Node
from office_mcp.shared.results import SendStatus

SUBJECT = "Project Atlas budget"
BODY = "Hi Bob,\nThe numbers are attached."
CLOSING_BODY = BODY + "\n\nBest regards,"
DISCARDED = "a copy is in Deleted Items"


def new_mail(
    to: tuple[str, ...] = ("bob@example.com",),
    body: str = BODY,
    allow_duplicate: bool = False,
    dry_run: bool = False,
) -> NewMail:
    return NewMail(
        to=to, subject=SUBJECT, body=body, allow_duplicate=allow_duplicate, dry_run=dry_run
    )


def fixed_review(body: str, subject: str = SUBJECT) -> Review:
    draft = Draft(channel=Channel.EMAIL, body=body, subject=subject)
    return Review(
        draft=draft,
        violations=(),
        fixes=(Fix("signoff.appended", "added the closing"),),
        persona_version="v1",
    )


def violation_review() -> Review:
    draft = Draft(channel=Channel.EMAIL, body=BODY, subject=SUBJECT)
    broken = Violation("characters", "forbidden character U+2014 EM DASH", Severity.ERROR)
    return Review(draft=draft, violations=(broken,), fixes=(), persona_version="v1")


def test_new_mail_inserts_the_reviewed_text_and_confirms_from_sent_items() -> None:
    rig = compose_rig(guard=RecordingGuard([fixed_review(CLOSING_BODY)]))
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)

    result = rig.service.send(new_mail(), rig.deadline())

    record = result.to_record()
    assert result.status is SendStatus.SENT
    assert result.verified is True
    assert result.message_ref
    assert record["text_sent"] == CLOSING_BODY
    assert record["persona"]["fixes"][0]["rule"] == "signoff.appended"
    assert rig.compose.send_presses == 1
    assert rig.compose.calls == [
        "open_new",
        "dismiss_popovers",
        "add:To:bob@example.com",
        "set_subject",
        "insert_body",
        "press_send",
    ]
    intended, body_read, subject_read = rig.guard.verified[0]
    assert intended.body == CLOSING_BODY
    assert body_read == CLOSING_BODY
    assert subject_read == SUBJECT
    assert rig.scope_log == ["enter foreground=True", "ensure_front", "exit"]


def test_persona_violation_makes_no_ui_call() -> None:
    rig = compose_rig(guard=RecordingGuard([violation_review()]))

    with pytest.raises(PersonaViolationError):
        rig.service.send(new_mail(), rig.deadline())

    assert rig.compose.calls == []
    assert rig.outlook.visible_calls() == []
    assert rig.scope_log == []


def test_allowlist_rejection_makes_no_ui_call() -> None:
    rig = compose_rig(allowlist=OWNER)

    with pytest.raises(RecipientNotAllowedError):
        rig.service.send(new_mail(to=("nobody@example.invalid",)), rig.deadline())

    assert rig.compose.calls == []
    assert rig.outlook.visible_calls() == []
    assert rig.scope_log == []


def test_allowlist_rejection_on_reply_all_happens_before_any_ui() -> None:
    latest = latest_message(
        inbox_key(),
        to=(Person("Jane Doe", OWNER), Person("Bob Sample", "bob@contoso.com")),
    )
    rig = compose_rig(latest=latest, allowlist="@example.com")

    with pytest.raises(RecipientNotAllowedError, match=r"bob@contoso\.com"):
        rig.service.reply(ReplyMail(key=inbox_key(), body=BODY, reply_all=True), rig.deadline())

    assert rig.compose.calls == []
    assert rig.outlook.visible_calls() == []


def test_invalid_address_is_refused_before_anything() -> None:
    rig = compose_rig()

    with pytest.raises(ComposeError, match="outlook_people"):
        rig.service.send(new_mail(to=("Bob Sample",)), rig.deadline())

    assert rig.guard.reviewed == []
    assert rig.compose.calls == []


def test_dry_run_send_makes_no_ui_call_and_returns_the_final_text() -> None:
    rig = compose_rig(guard=RecordingGuard([fixed_review(CLOSING_BODY)]))

    result = rig.service.send(new_mail(dry_run=True), rig.deadline())

    assert result.status is SendStatus.DRY_RUN
    assert result.to_record()["text_sent"] == CLOSING_BODY
    assert result.to_record()["ok"] is True
    assert rig.compose.calls == []
    assert rig.outlook.visible_calls() == []
    assert rig.reader.calls == []
    assert rig.scope_log == []


def test_dry_run_reply_reads_only_the_latest_message() -> None:
    rig = compose_rig(guard=RecordingGuard([fixed_review(CLOSING_BODY, subject="")]))

    result = rig.service.reply(
        ReplyMail(key=inbox_key(), body=BODY, dry_run=True, folder="Inbox"), rig.deadline()
    )

    record = result.to_record()
    assert result.status is SendStatus.DRY_RUN
    assert record["text_sent"] == CLOSING_BODY
    assert record["subject"] == "RE: Project Atlas budget"
    assert record["in_reply_to"]["key"] == inbox_key()
    assert record["recipients"]["to"] == ["omar@example.com"]
    assert rig.reader.calls == [(inbox_key(), "Inbox")]
    assert rig.compose.calls == []
    assert rig.outlook.visible_calls() == []
    reviewed = rig.guard.reviewed[0]
    assert reviewed.context == "Could you send the numbers?"
    assert reviewed.context_author == "Omar Haddad"


def test_dry_run_forward_reads_only_the_latest_message() -> None:
    rig = compose_rig()

    result = rig.service.forward(
        ForwardMail(key=inbox_key(), to=("bob@example.com",), body="FYI", dry_run=True),
        rig.deadline(),
    )

    assert result.status is SendStatus.DRY_RUN
    assert result.to_record()["text_sent"] == "FYI"
    assert result.to_record()["forwarded"]["key"] == inbox_key()
    assert rig.reader.calls == [(inbox_key(), None)]
    assert rig.compose.calls == []
    assert rig.outlook.visible_calls() == []


def test_reply_refuses_a_non_key_before_reading() -> None:
    rig = compose_rig()

    with pytest.raises(TargetError):
        rig.service.reply(ReplyMail(key="Project Atlas", body=BODY), rig.deadline())

    assert rig.reader.calls == []


@pytest.mark.parametrize(
    ("compose", "message"),
    [
        (FakeCompose(unreadable={"bob@example.com"}), "could not read the address"),
        (
            FakeCompose(
                extra_tokens={
                    RecipientField.TO: [RecipientToken("Eve", "eve@example.com")],
                }
            ),
            "shows 2 recipients",
        ),
        (
            FakeCompose(
                extra_tokens={RecipientField.CC: [RecipientToken("Eve", "eve@example.com")]}
            ),
            "Cc shows 1 recipients",
        ),
        (FakeCompose(pending={RecipientField.TO: "bob@exa"}), "not a recipient"),
    ],
)
def test_token_problems_discard_and_raise(compose: FakeCompose, message: str) -> None:
    rig = compose_rig(compose=compose)

    with pytest.raises(ComposeError, match=message) as caught:
        rig.service.send(new_mail(), rig.deadline())

    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1
    assert rig.compose.open == []
    assert DISCARDED in error_text(caught.value)


def test_token_with_a_different_address_discards() -> None:
    rig = compose_rig(compose=FakeCompose(shown_as={"bob@example.com": "bob@contoso.com"}))

    with pytest.raises(ComposeError, match="recipients differ"):
        rig.service.send(new_mail(), rig.deadline())

    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1


def test_reply_subject_mismatch_discards() -> None:
    rig = compose_rig(
        compose=FakeCompose(
            reply_subject="RE: Something else", prefilled={RecipientField.TO: ["omar@example.com"]}
        )
    )

    with pytest.raises(ComposeError, match="does not match") as caught:
        rig.service.reply(ReplyMail(key=inbox_key(), body=BODY), rig.deadline())

    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1
    assert DISCARDED in error_text(caught.value)


def test_reply_sends_once_to_the_expected_sender() -> None:
    compose = FakeCompose(prefilled={RecipientField.TO: ["omar@example.com"]})
    rig = compose_rig(compose=compose)
    rig.compose.on_send = lambda: rig.deliver_to_sent("RE: Project Atlas budget")

    result = rig.service.reply(ReplyMail(key=inbox_key(), body=BODY), rig.deadline())

    assert result.status is SendStatus.SENT
    assert result.subject == "RE: Project Atlas budget"
    assert rig.compose.send_presses == 1
    assert rig.compose.calls[0] == "open_reply:inbox-row:Reply"
    assert "insert_body" in rig.compose.calls
    assert "set_subject" not in rig.compose.calls


def test_reply_is_confirmed_by_the_conversation_topic_in_sent_items() -> None:
    compose = FakeCompose(prefilled={RecipientField.TO: ["omar@example.com"]})
    rig = compose_rig(compose=compose)
    rig.compose.on_send = lambda: rig.deliver_to_sent("Project Atlas budget")

    result = rig.service.reply(ReplyMail(key=inbox_key(), body=BODY), rig.deadline())

    assert result.status is SendStatus.SENT
    assert result.verified is True


def test_reply_all_excludes_the_owner_and_compares_to_and_cc_together() -> None:
    latest = latest_message(
        inbox_key(),
        to=(Person("Jane Doe", OWNER), Person("Bob Sample", "bob@example.com")),
        cc=(Person("Alex Example", "alex@example.com"),),
    )
    compose = FakeCompose(
        prefilled={
            RecipientField.TO: ["omar@example.com", "bob@example.com"],
            RecipientField.CC: ["alex@example.com"],
        }
    )
    rig = compose_rig(compose=compose, latest=latest)

    result = rig.service.reply(
        ReplyMail(key=inbox_key(), body=BODY, reply_all=True), rig.deadline()
    )

    assert result.to_record()["recipients"]["to"] == [
        "omar@example.com",
        "bob@example.com",
        "alex@example.com",
    ]
    assert OWNER not in result.targets
    assert rig.compose.calls[0] == "open_reply:inbox-row:Reply all"
    assert rig.compose.send_presses == 1


def test_reply_all_to_a_self_only_mail_goes_to_the_owner() -> None:
    latest = latest_message(inbox_key(), sender=Person("Jane Doe", OWNER))

    rig = compose_rig(latest=latest)
    result = rig.service.reply(
        ReplyMail(key=inbox_key(), body=BODY, reply_all=True, dry_run=True), rig.deadline()
    )

    assert result.targets == (OWNER,)


def test_reply_all_with_hidden_recipients_is_refused_before_ui() -> None:
    rig = compose_rig(latest=latest_message(inbox_key(), hidden=2))

    with pytest.raises(ComposeError, match="hides 2 recipients"):
        rig.service.reply(ReplyMail(key=inbox_key(), body=BODY, reply_all=True), rig.deadline())

    assert rig.compose.calls == []


def test_busy_body_raises_and_inserts_nothing() -> None:
    rig = compose_rig(compose=FakeCompose(baseline="Unfinished thoughts\n\nJane Doe"))

    with pytest.raises(ComposeBusyError) as caught:
        rig.service.send(new_mail(), rig.deadline())

    assert "insert_body" not in rig.compose.calls
    assert not any(call.startswith("add:") for call in rig.compose.calls)
    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1
    assert DISCARDED in error_text(caught.value)


def test_busy_reply_is_left_open_untouched() -> None:
    compose = FakeCompose(
        reply_baseline="My unfinished reply\n\nFrom: Omar Haddad",
        prefilled={RecipientField.TO: ["omar@example.com"]},
    )
    rig = compose_rig(compose=compose)

    with pytest.raises(ComposeBusyError) as caught:
        rig.service.reply(ReplyMail(key=inbox_key(), body=BODY), rig.deadline())

    assert rig.compose.discards == 0
    assert len(rig.compose.open) == 1
    assert "insert_body" not in rig.compose.calls
    assert "left open" in error_text(caught.value)


def test_substituted_dash_discards_without_sending() -> None:
    compose = FakeCompose(body_transform=lambda text: text.replace(" - ", " \u2013 "))
    rig = compose_rig(compose=compose)

    with pytest.raises(TextEntryError, match=r"\[U\+2013\]") as caught:
        rig.service.send(new_mail(body="Plan A - then plan B"), rig.deadline())

    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1
    assert DISCARDED in error_text(caught.value)


def test_alert_after_send_is_refused_and_discarded() -> None:
    alert = AlertSheet(
        text="Choose a sensitivity label",
        buttons=("Send Anyway", "Cancel"),
        node=Node(role="AXSheet"),
    )
    rig = compose_rig(compose=FakeCompose(alert_after_send=alert))

    with pytest.raises(SendRefusedError, match="nothing was sent") as caught:
        rig.service.send(new_mail(), rig.deadline())

    assert rig.compose.send_presses == 1
    assert "alert:Cancel" in rig.compose.calls
    assert "alert:Send Anyway" not in rig.compose.calls
    assert rig.compose.discards == 1
    assert DISCARDED in error_text(caught.value)
    assert rig.sends.last(60) is None


def test_missing_sent_row_returns_unconfirmed_without_raising() -> None:
    rig = compose_rig()

    result = rig.service.send(new_mail(), rig.deadline())

    record = result.to_record()
    assert result.status is SendStatus.UNCONFIRMED
    assert record["ok"] is False
    assert record["verified"] is False
    assert any("check Sent Items" in note for note in record["notes"])
    assert rig.compose.send_presses == 1
    assert rig.compose.discards == 0


def test_compose_still_open_after_send_is_unconfirmed_and_not_discarded() -> None:
    rig = compose_rig(compose=FakeCompose(close_on_send=False))

    result = rig.service.send(new_mail(), rig.deadline())

    assert result.status is SendStatus.UNCONFIRMED
    assert "the compose window was still open after Send" in result.notes
    assert rig.compose.discards == 0
    assert rig.compose.send_presses == 1


def test_press_error_after_send_is_reported_not_raised() -> None:
    rig = compose_rig(compose=FakeCompose(send_error=AxActionError("AXPress failed (-25200)")))
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)

    result = rig.service.send(new_mail(), rig.deadline())

    assert result.status is SendStatus.SENT
    assert any("pressing Send reported" in note for note in result.notes)
    assert rig.compose.send_presses == 1
    assert rig.compose.discards == 0


def test_cleanup_failure_after_send_becomes_a_note() -> None:
    rig = compose_rig(scope=failing_exit_factory("could not put the Outlook view back"))
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)

    result = rig.service.send(new_mail(), rig.deadline())

    assert result.status is SendStatus.SENT
    assert any("could not put the Outlook view back" in note for note in result.notes)


def test_memory_duplicate_guard_refuses_the_same_mail_before_ui() -> None:
    rig = compose_rig()
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)
    rig.service.send(new_mail(), rig.deadline())
    calls_before = list(rig.compose.calls)

    with pytest.raises(DuplicateSendError):
        rig.service.send(new_mail(), rig.deadline())

    assert rig.compose.calls == calls_before
    assert rig.compose.send_presses == 1


def test_memory_duplicate_guard_expires_after_two_minutes() -> None:
    rig = compose_rig()
    rig.service.send(new_mail(), rig.deadline())
    rig.clock.advance(121)

    result = rig.service.send(new_mail(allow_duplicate=False), rig.deadline())

    assert result.status is SendStatus.UNCONFIRMED
    assert rig.compose.send_presses == 2


def test_sent_items_duplicate_refuses_before_opening_compose() -> None:
    rig = compose_rig(sent_rows=[sent_row(SUBJECT)])

    with pytest.raises(DuplicateSendError, match="last 10 minutes"):
        rig.service.send(new_mail(), rig.deadline())

    assert rig.compose.calls == []


def test_old_sent_row_is_not_a_duplicate() -> None:
    rig = compose_rig(sent_rows=[sent_row(SUBJECT, time_label="8:00 AM")])

    result = rig.service.send(new_mail(), rig.deadline())

    assert rig.compose.send_presses == 1
    assert result.status is SendStatus.UNCONFIRMED


def test_allow_duplicate_skips_both_layers_and_needs_a_new_sent_row() -> None:
    rig = compose_rig(sent_rows=[sent_row(SUBJECT)])
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)

    first = rig.service.send(new_mail(allow_duplicate=True), rig.deadline())
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT, time_label="9:29 AM")
    second = rig.service.send(new_mail(allow_duplicate=True), rig.deadline())

    assert first.status is SendStatus.SENT
    assert second.status is SendStatus.SENT
    assert first.message_ref != second.message_ref
    assert rig.compose.send_presses == 2


def test_existing_sent_row_alone_does_not_confirm_a_send() -> None:
    rig = compose_rig(sent_rows=[sent_row(SUBJECT)])

    result = rig.service.send(new_mail(allow_duplicate=True), rig.deadline())

    assert result.status is SendStatus.UNCONFIRMED


def test_out_of_time_stops_before_send() -> None:
    rig = compose_rig()

    with pytest.raises(UiTimeoutError) as caught:
        rig.service.send(new_mail(), rig.deadline(seconds=1))

    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1
    assert DISCARDED in error_text(caught.value)


def test_forward_with_empty_body_types_recipients_only() -> None:
    compose = FakeCompose(reply_subject="FW: Project Atlas budget")
    rig = compose_rig(compose=compose)

    result = rig.service.forward(
        ForwardMail(key=inbox_key(), to=("bob@example.com",)), rig.deadline()
    )

    assert rig.compose.calls[0] == "open_reply:inbox-row:Forward"
    assert "insert_body" not in rig.compose.calls
    assert "add:To:bob@example.com" in rig.compose.calls
    assert rig.compose.send_presses == 1
    assert result.to_record()["text_sent"] == ""
    assert rig.guard.verified == []


def test_forward_with_reply_prefix_is_refused() -> None:
    rig = compose_rig(compose=FakeCompose(reply_subject="RE: Project Atlas budget"))

    with pytest.raises(ComposeError, match="does not match"):
        rig.service.forward(ForwardMail(key=inbox_key(), to=("bob@example.com",)), rig.deadline())

    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1


def test_sent_folder_is_visited_for_the_baseline() -> None:
    rig = compose_rig()
    rig.service.send(new_mail(), rig.deadline())

    assert "select:folder-sent" in rig.outlook.visible_calls()
    assert rig.outlook.folder == SENT_FOLDER


def test_foreign_alert_after_a_closed_send_is_left_alone_and_blocks_a_retry() -> None:
    foreign = AlertSheet(
        text="Save changes to the other draft?",
        buttons=("Cancel", "Save"),
        node=Node(role="AXSheet"),
    )
    compose = FakeCompose(sheets=[("other-window", foreign), ("compose1", foreign)])
    rig = compose_rig(compose=compose)
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)

    result = rig.service.send(new_mail(), rig.deadline())

    assert result.status is SendStatus.SENT
    assert not any(call.startswith("alert:") for call in rig.compose.calls)
    assert rig.compose.discards == 0
    with pytest.raises(DuplicateSendError):
        rig.service.send(new_mail(), rig.deadline())
    assert rig.compose.send_presses == 1


def test_new_alert_on_another_window_is_not_a_refusal() -> None:
    foreign = AlertSheet(text="Other", buttons=("Cancel",), node=Node(role="AXSheet"))
    rig = compose_rig()

    def owner_alert() -> None:
        rig.compose.sheets.append(("other-window", foreign))

    rig.compose.on_send = owner_alert

    result = rig.service.send(new_mail(), rig.deadline())

    assert result.status is SendStatus.UNCONFIRMED
    assert "alert:Cancel" not in rig.compose.calls
    assert rig.sends.last(60) is not None


def test_unexpected_error_after_pressing_send_still_records_the_send() -> None:
    rig = compose_rig(compose=FakeCompose(send_error=RuntimeError("bridge failed")))

    with pytest.raises(RuntimeError):
        rig.service.send(new_mail(), rig.deadline())

    with pytest.raises(DuplicateSendError):
        rig.service.send(new_mail(), rig.deadline())
    assert rig.compose.send_presses == 1


LABEL = ("Confidential", "Internal")


def test_sensitivity_label_is_set_before_send_and_reported() -> None:
    rig = compose_rig()
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)
    mail = NewMail(to=("bob@example.com",), subject=SUBJECT, body=BODY, sensitivity=LABEL)

    result = rig.service.send(mail, rig.deadline())

    assert result.status is SendStatus.SENT
    assert result.to_record()["sensitivity"] == list(LABEL)
    calls = rig.compose.calls
    assert calls.index("sensitivity:Confidential/Internal") < calls.index("press_send")
    assert calls.index("insert_body") < calls.index("sensitivity:Confidential/Internal")


def test_sensitivity_label_that_does_not_show_discards_without_sending() -> None:
    rig = compose_rig(compose=FakeCompose(label_shown_as={"Internal": "General"}))
    mail = NewMail(to=("bob@example.com",), subject=SUBJECT, body=BODY, sensitivity=LABEL)

    with pytest.raises(TextEntryError, match="Internal") as caught:
        rig.service.send(mail, rig.deadline())

    assert rig.compose.send_presses == 0
    assert rig.compose.discards == 1
    assert DISCARDED in error_text(caught.value)


@pytest.mark.parametrize("path", [("A", "B", "C", "D"), ("Confidential", " "), ("Two\nlines",)])
def test_bad_sensitivity_path_is_refused_before_any_ui(path: tuple[str, ...]) -> None:
    rig = compose_rig()
    mail = NewMail(to=("bob@example.com",), subject=SUBJECT, body=BODY, sensitivity=path)

    with pytest.raises(ComposeError, match="sensitivity"):
        rig.service.send(mail, rig.deadline())

    assert rig.compose.calls == []
    assert rig.scope_log == []


def test_dry_run_reports_the_sensitivity_without_opening_anything() -> None:
    rig = compose_rig()
    mail = NewMail(
        to=("bob@example.com",), subject=SUBJECT, body=BODY, sensitivity=("General",), dry_run=True
    )

    result = rig.service.send(mail, rig.deadline())

    assert result.status is SendStatus.DRY_RUN
    assert result.to_record()["sensitivity"] == ["General"]
    assert rig.compose.calls == []


def test_reply_takes_a_sensitivity_label() -> None:
    rig = compose_rig(compose=FakeCompose(prefilled={RecipientField.TO: ["omar@example.com"]}))

    rig.service.reply(
        ReplyMail(key=inbox_key(), body=BODY, sensitivity=("General",)), rig.deadline()
    )

    assert "sensitivity:General" in rig.compose.calls


def test_forward_takes_a_sensitivity_label() -> None:
    rig = compose_rig(compose=FakeCompose(reply_subject="FW: Project Atlas budget"))

    rig.service.forward(
        ForwardMail(key=inbox_key(), to=("bob@example.com",), sensitivity=("General",)),
        rig.deadline(),
    )

    assert "sensitivity:General" in rig.compose.calls
