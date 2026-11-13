# Security policy

## Reporting a vulnerability

Report vulnerabilities privately through GitHub's [private vulnerability reporting](https://github.com/MajidRaimi/mcp-suite/security/advisories/new). Do not open a public issue or pull request for a security problem.

Include the affected server, the version or commit, steps to reproduce, and the impact you expect. Reports are handled on a best-effort basis. Once a fix is released, the advisory is published with credit unless you ask otherwise.

## Supported versions

| Version | Supported |
|---|---|
| `main` | Yes |
| anything else | No |

## Threat model

These servers connect personal accounts and desktop apps to an LLM. They assume the machine they run on and the person who starts them are trusted. Everything else is not.

**Network exposure**

- HTTP servers bind to `127.0.0.1` and have no authentication by default. Anyone who can reach the port can call every tool.
- Set `<NAME>_MCP_TOKEN` for any access beyond your own user on the same machine. Every HTTP request, including `/health`, then needs `Authorization: Bearer <token>`, checked in constant time.
- For remote use, go through an SSH tunnel, a private VPN, or a reverse proxy with TLS. Never bind to `0.0.0.0` on a reachable interface.
- On loopback the MCP SDK checks the `Host` and `Origin` headers by default to block DNS rebinding. `<NAME>_MCP_ALLOWED_HOSTS` adds names to that list and never removes the check. When the server is bound to any other address, the SDK skips the check unless `<NAME>_MCP_ALLOWED_HOSTS` is set, so always set both `<NAME>_MCP_ALLOWED_HOSTS` and `<NAME>_MCP_TOKEN` when binding to a Docker bridge or LAN address.

**Untrusted content and prompt injection**

- Messages, mail, chat names and contact names are written by third parties and returned to the model as tool output. Treat them as hostile input that can try to steer the model.
- Any tool that sends or changes data uses a draft step and a separate send step. The rule that the model sends only on your explicit instruction is enforced by tool descriptions, so keep your client's approval prompt on for send tools.

**whatsapp**

- WAHA is an unofficial WhatsApp Web client. Using it may violate WhatsApp's Terms of Service and can get the number banned.
- The WAHA API key grants full control of the paired account. Keep it out of shell history and version control, and keep WAHA itself bound to loopback.
- Attachments are read from the server host and only from inside `WHATSAPP_MCP_ATTACHMENTS_DIR`. File attachments are disabled until it is set, and a path outside that folder is rejected.
- Drafts are stored on disk (mode 600) and are removed when sent or, once expired, at the next draft or send.

**office**

- The macOS Accessibility permission lets the process read and control every app, not only Outlook and Teams. Granting it to an SSH helper grants it to every SSH session.
- The tools change what is on screen, and `teams_read` marks chats as read in Teams.
- Use a dedicated account or VM where you can.

**Data leaving the machine**

- Tool output goes to whichever LLM provider your client uses. Choose one you trust with your messages and mail, or use a local model.

## Out of scope

- Behaviour of WAHA, WhatsApp, Microsoft Outlook, Microsoft Teams or MCP clients themselves.
- Attacks that need an already compromised user account on the host.
- Running a server bound to a public interface without a token, against the documentation.
