from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC

from fakes.clock import FakeClock
from fakes.drivers import FAKE_TODAY, NO_IDLE_WAIT
from fakes.events import RecordingEvents
from fakes.guard import RecordingGuard
from fakes.p5_world import OWNER, FakeChat, TeamsWorld, WorldOptions, other, own
from fakes.trees import AppRig, app_rig, window
from fakes.workspace import FakeWorkspace
from office_mcp.features.teams.chats.opener import ChatOpener
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.driver import TeamsDriver
from office_mcp.features.teams.writing.compose import ComposeBox
from office_mcp.features.teams.writing.factory import assemble_writing, writing_context
from office_mcp.features.teams.writing.service import TeamsWritingService
from office_mcp.shared.ax import AccessMode
from office_mcp.shared.recipients import RecipientGuard, parse_allowlist
from office_mcp.shared.runner import Deadline
from office_mcp.shared.text_input import TextWriter

SELF_THREAD = "48:notes"
OWNER_APP = 100


@dataclass
class WritingRig:
    service: TeamsWritingService
    world: TeamsWorld
    guard: RecordingGuard
    events: RecordingEvents
    workspace: FakeWorkspace
    app: AppRig

    @property
    def clock(self) -> FakeClock:
        return self.app.clock

    def deadline(self, seconds: float = 45.0) -> Deadline:
        return Deadline.after(self.clock, seconds)


def writing_rig(
    world: TeamsWorld,
    allowlist: str | None = None,
    guard: RecordingGuard | None = None,
    max_members: int = 25,
) -> WritingRig:
    rig = app_rig(window("Teams", *(chat.box for chat in world.chats)), mode=AccessMode.WEB)
    world.install_box_hooks()
    events = RecordingEvents()
    events.on_press = world.on_event
    desk = FakeWorkspace(frontmost=OWNER_APP)
    desk.running.add(rig.app.require_pid())
    waiter = rig.app.waiter
    chats = ChatService(world, lambda: FAKE_TODAY, waiter)
    driver = TeamsDriver(
        app=rig.app,
        gateway=world,
        chats=chats,
        opener=ChatOpener(world, chats, waiter),
        text=TextWriter(rig.app, events, world.pasteboard, waiter),
        workspace=desk,
        waiter=waiter,
        zone=UTC,
        focus=NO_IDLE_WAIT,
    )
    recording = guard or RecordingGuard()
    service = assemble_writing(
        context=writing_context(driver, RecipientGuard(parse_allowlist(allowlist), SELF_THREAD)),
        box=ComposeBox(world, driver.text),
        guard=recording,
        clock=rig.clock,
        max_members=max_members,
    )
    return WritingRig(service, world, recording, events, desk, rig)


ALICE = "19:alice_alex@unq.gbl.spaces"
BOB = "19:bob_alex@unq.gbl.spaces"
GROUP = "19:atlasteam@thread.v2"
BIG_GROUP = "19:allhands@thread.v2"
MEETING = "19:meeting_weeklysync@thread.v2"
CHANNEL = "19:general@thread.tacv2"
BOB_QUESTION = "1790000001000"
MY_ANSWER = "1790000002000"
SELF_NOTE = "1790000003000"
ALICE_NOTE = "1790000004000"


def standard_chats(bob_draft: str = "") -> list[FakeChat]:
    return [
        FakeChat(
            ALICE,
            "Alice Example",
            participants=["Alice Example", OWNER],
            messages=[other(int(ALICE_NOTE), "Alice Example", "See you at noon")],
        ),
        FakeChat(
            BOB,
            "Bob Sample",
            participants=["Bob Sample", OWNER],
            messages=[
                other(int(BOB_QUESTION), "Bob Sample", "Can you review the Project Atlas budget?"),
                own(int(MY_ANSWER), "Sure, looking now"),
            ],
            draft=bob_draft,
        ),
        FakeChat(
            SELF_THREAD,
            "Alex Example (You)",
            participants=[OWNER],
            messages=[own(int(SELF_NOTE), "Remember the Atlas budget")],
        ),
        FakeChat(
            GROUP,
            "Project Atlas team",
            kind="Group chat",
            participants=["Jane Doe", "Omar Haddad", "Bob Sample", OWNER],
            messages=[other(1790000000500, "Jane Doe", "Agenda is ready")],
        ),
        FakeChat(
            BIG_GROUP,
            "Contoso all hands",
            kind="Group chat",
            participants=[f"Person {number}" for number in range(30)],
            messages=[other(1790000000600, "Jane Doe", "Welcome everyone")],
        ),
        FakeChat(
            MEETING,
            "Weekly sync",
            kind="Meeting chat",
            participants=["Jane Doe", OWNER],
            messages=[other(1790000000700, "Jane Doe", "Notes are posted")],
        ),
    ]


def standard_world(options: WorldOptions | None = None, bob_draft: str = "") -> TeamsWorld:
    return TeamsWorld(standard_chats(bob_draft), open_thread=ALICE, options=options)
