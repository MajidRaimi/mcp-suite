# __name__-mcp

One sentence on what this server connects to and what it lets an MCP client do.

| Platform | Access | Transports | Default port |
|---|---|---|---|
| Linux, macOS | read-only | stdio, streamable HTTP | 8770 |

## What it does

| Tool | Purpose | Key parameters | Side effects |
|---|---|---|---|
| `__name___status` | Shows that the server runs, whose account it serves and the local time | none | none |

Replace this table with every tool the server exposes. List the limits (page sizes, how many upstream calls a tool can make) and every side effect, such as marking items read or sending anything.

## How it works

Describe the upstream system (a REST API, a local app, a file format), how the server authenticates to it, and what it never does.

## Requirements

- Python 3.13 and [uv](https://docs.astral.sh/uv/)
- Anything else the upstream system needs (Docker, an account, an OS permission)

## Setup

From the repository root:

```sh
uv sync --all-packages
cp servers/__name__/.env.example servers/__name__/.env
```

Fill in `.env`. The server does not load `.env` by itself. Run it with `uv run --env-file servers/__name__/.env --package __name__-mcp __name__-mcp`, or copy the values into your client's `env` block.

## Configuration

| Variable | Default | Meaning |
|---|---|---|
| `__NAME___MCP_TRANSPORT` | `stdio` | `stdio` or `streamable-http` |
| `__NAME___MCP_HOST` | `127.0.0.1` | HTTP bind address |
| `__NAME___MCP_PORT` | `8770` | HTTP port |
| `__NAME___MCP_TOKEN` | empty | When set, every HTTP request needs `Authorization: Bearer <token>` |
| `__NAME___MCP_ALLOWED_HOSTS` | empty | Extra `host[:port]` values accepted in the `Host` header, comma separated |
| `__NAME___MCP_TIMEZONE` | system zone | IANA zone used for every displayed time |
| `__NAME___MCP_OWNER_NAME` | `the user` | Name used in tool descriptions |

CLI flags `--transport`, `--host` and `--port` override the environment.

## Running

Over stdio your client normally starts the server itself (see below). For clients that connect by URL, start it over HTTP:

```sh
uv run --env-file servers/__name__/.env --package __name__-mcp __name__-mcp --transport streamable-http
```

In a second terminal:

```sh
curl http://127.0.0.1:8770/health
```

## Connecting clients

```sh
claude mcp add --transport http --scope user __name__ http://127.0.0.1:8770/mcp
claude mcp add __name__ -- uv run --project /path/to/mcp-suite/servers/__name__ __name__-mcp
```

See the root README for Claude Desktop and Open WebUI.

## Security and limitations

State what the tools can read, what they can change, which actions need draft-then-confirm, and what data leaves the machine.

## Development

```sh
cd servers/__name__
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pytest
```
