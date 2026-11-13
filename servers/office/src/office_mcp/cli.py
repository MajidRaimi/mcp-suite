from __future__ import annotations

import argparse
import io
import json
import sys
from collections.abc import Callable
from typing import Any

from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.server import build_tools
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.settings import Settings
from office_mcp.tools import OfficeTools


def commands(tools: OfficeTools) -> dict[str, Callable[..., Any]]:
    return {
        "status": tools.office_status,
        "folders": tools.outlook_folders,
        "mail": tools.outlook_conversations,
        "read-mail": tools.outlook_read,
        "chats": tools.teams_chats,
        "read-chat": tools.teams_read,
        "needs-reply": tools.teams_needs_reply,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="office", description="Debug the office MCP tools")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("status")
    commands.add_parser("folders")
    mail = commands.add_parser("mail")
    mail.add_argument("--folder", default="Inbox")
    mail.add_argument("--unread-only", action="store_true")
    mail.add_argument("--days", type=int)
    mail.add_argument("--limit", type=int, default=50)
    mail.add_argument("--load-more", type=int, default=0)
    read_mail = commands.add_parser("read-mail")
    read_mail.add_argument("ref")
    read_mail.add_argument("--folder")
    read_mail.add_argument("--mark-read", dest="keep_unread", action="store_false")
    read_mail.add_argument("--max-body", type=int, default=20000)
    chats = commands.add_parser("chats")
    chats.add_argument("--unread-only", action="store_true")
    chats.add_argument("--section")
    chats.add_argument("--limit", type=int, default=0)
    read_chat = commands.add_parser("read-chat")
    read_chat.add_argument("chat")
    read_chat.add_argument("--limit", type=int, default=30)
    read_chat.add_argument("--older", type=int, default=0)
    waiting = commands.add_parser("needs-reply")
    waiting.add_argument("--days", type=int, default=7)
    waiting.add_argument("--include-meetings", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> None:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(errors="backslashreplace")
    args = vars(build_parser().parse_args(argv))
    name = args.pop("command")
    try:
        command = commands(build_tools(Settings.from_env()))[name]
        result = command(**args)
    except (OfficeError, ToolError, OSError) as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False))
        raise SystemExit(1) from error
    print(json.dumps(result, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
