# Contributing

Thanks for helping. This file covers how to set up, what the checks are, and the few house rules that CI and reviewers enforce.

## Setup

You need [git](https://git-scm.com) and [uv](https://docs.astral.sh/uv/). uv installs Python 3.13 for you.

```sh
git clone https://github.com/MajidRaimi/mcp-suite
cd mcp-suite
uv sync --all-packages
```

Use uv for everything: `uv add` inside a server directory to add a dependency, `uv run` to run tools. Do not use pip, and do not commit a second lockfile.

## Checks

Run these in the directory of every server you touched before you push:

```sh
uv run ruff check .
uv run ruff format --check .
uv run mypy src tests
uv run pytest
```

From the repository root, also run:

```sh
uv lock --check
```

CI runs the same commands for every server in `servers/`, on `ubuntu-latest` plus any runners listed in `servers/<name>/ci.json`. It also scaffolds a server from `templates/server` and checks it, scans for secrets with gitleaks, rejects em dashes, and checks that every server has `README.md`, `.env.example`, `pyproject.toml` and `tests/`.

## Commits and pull requests

- Use [Conventional Commits](https://www.conventionalcommits.org): `type(scope): description`.
- The scope is the server name for server changes (`feat(whatsapp): add stdio transport`, `fix(office): restore unread state`), or `repo`, `ci` or `docs` for shared files.
- Keep a PR to one server or one shared concern when you can.
- Fill in the PR template checklist.

## Adding a server

Follow [Adding a new server](README.md#adding-a-new-server) in the README. The short version:

```sh
uv run --no-project python templates/scaffold.py <name> --port <next free port>
uv sync --all-packages
```

Then implement tools under `features/`, keep every item of the [server contract](README.md#server-contract), add the catalog row, and open a PR.

## House rules

- **Strict typing.** Every server passes `mypy --strict`. Avoid `Any`. Type every parameter and return value.
- **Feature-first layout.** Code lives in `features/<domain>/` and `shared/`, never in type buckets such as `utils/` or `models/` at the package root.
- **Small units, custom errors.** Keep functions short and single-purpose. Raise errors that subclass the server's base error in `shared/errors.py`, and never swallow exceptions silently.
- **No code comments.** Do not write comments or docstrings in any file: Python, YAML, TOML, shell or anything else. Pick names that make the code read on its own. If something needs explaining, it belongs in the README.
- **No em dashes.** Do not use the em dash character anywhere in the repository, including prose. Use a period, comma, colon or parentheses instead. CI fails on it.
- **No personal data or secrets.** Code, tests, fixtures and docs use invented names, numbers and addresses only (for example `Alice Example` and `+1 555 0100`). `.env.example` files hold placeholders, never real values.
- **Draft, then confirm.** Any tool that sends, writes or deletes uses a two-step flow and says so in its description.

## Why a uv workspace and not Turborepo

Every server here is pure Python. A JavaScript monorepo tool would add Node, a package manager and per-server shim `package.json` files only to call `uv run`, with no real build graph to cache, because the checks already take seconds. The people who use this repo are mostly self-hosters, and one toolchain (uv) is easier for them to install and trust. A uv workspace gives one lockfile, one environment and automatic discovery of new servers through the `servers/*` glob, while each server stays installable on its own.

## Security issues

Do not open a public issue for a vulnerability. See [SECURITY.md](SECURITY.md).
