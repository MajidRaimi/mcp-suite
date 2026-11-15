# Persona

The persona tells the model how the owner writes and holds hard rules that the server checks before anything is sent. It is a Markdown file that lives outside the repository.

## Location and loading

- Path: `OFFICE_MCP_PERSONA_FILE`, default `~/.config/office-mcp/persona.md`. Never put it inside the repository; `.gitignore` lists `persona.md`, `persona-*.md`, `persona-corpus*.jsonl` and `persona-stats.md` as a backstop.
- `OFFICE_MCP_PERSONA_REQUIRED=true` blocks every send when the file is missing. Set it, together with an explicit `OFFICE_MCP_PERSONA_FILE`, on any deploy that sends.
- Keep the file private: `chmod 600`. `office_status` and `office_persona` report `permissions_ok=false` when group or world bits are set.
- The file is checked on every send. When its modification time or size changes it is parsed again, so edits apply without a restart. The version is the first 12 hex characters of the SHA-256 of the file and is returned with every send result.

| State | When | Effect |
|---|---|---|
| `loaded` | the file parsed cleanly | rules are enforced |
| `none` | the default path is missing and the persona is not required | no rules, empty prose, sends allowed |
| `missing` | an explicit path is missing, or the persona is required and missing | every send fails with `PersonaMissingError` |
| `invalid` | the file cannot be read, is not UTF-8, or breaks the schema | every send fails with `PersonaFormatError` naming the path, line and key |

Read tools never fail because of the persona. `office_persona` shows the state and the error.

## Format

An optional TOML front matter between a first line `+++` and the next `+++` line, then prose for the model. The schema is strict: an unknown key or a wrong type is an error that names the key path, for example `email.sign_off.fr`. A missing table means no rules of that kind. See `examples/persona.example.md` for a complete file.

| Key | Type | Meaning |
|---|---|---|
| `version` | int | must be 1 when present |
| `language.default` | `en` or `ar` | the language to write in (default `en`) |
| `language.mirror` | list of `en`, `ar` | reply in that language when the other side writes in it |
| `language.enforce` | bool | a mismatch is an error (default true) or only a warning |
| `forbid.character_sets` | list of `dashes` | `dashes` covers U+2012, U+2013, U+2014, U+2015, U+2212, U+2E3A, U+2E3B, U+FE31, U+FE32, U+FE58, U+FE63 and U+FF0D |
| `forbid.characters` | list of single characters | more forbidden characters |
| `forbid.emoji` | bool | refuse emoji (and, on Teams, emoticons such as `:)` and `<3`) |
| `forbid.emoji_allow` | list of single characters | characters exempt from the emoji rule; default copyright, registered and trade mark signs, which are still refused when followed by U+FE0F |
| `forbid.phrases` | list of strings | refused phrases, matched case-insensitively with Arabic diacritics, tatweel and letter forms folded |
| `forbid.patterns` | list of Python regular expressions | refused patterns, compiled when the file loads |
| `email`, `teams`, `rsvp` tables | | per channel rules below |
| `sign_off` | string, or table of `default`, `en`, `ar` | the closing line; `ar` is used for Arabic text when present, otherwise `default` |
| `sign_off_rule` | `required`, `forbidden` or `any` | default `any` |
| `nothing_after_sign_off` | bool | no line may follow the closing |
| `applies_to_accounts` | list of addresses or `@domain` | email only: the sign-off rules apply only when the sending account matches; empty means every account |
| `max_words`, `max_chars` | int | length warnings, never rejections |

## How a draft is checked

1. Line endings become LF, trailing spaces are removed and leading and trailing blank lines dropped.
2. Sign-off. With `required`, when the last non-empty line is already the closing nothing changes; when it looks like a different closing (for example "Thanks" or "Kind regards") the text is refused with `signoff.wrong`; otherwise the closing is appended after a blank line and reported as the fix `signoff.appended`. With `nothing_after_sign_off`, any line after the closing is refused with `signoff.trailing`. With `forbidden`, a closing-like last line is refused with `signoff.forbidden`. An email whose sending account is not in `applies_to_accounts` skips this step; an unknown account does not.
3. Errors: forbidden characters (each hit with an excerpt and the code point name), emoji, emoticons on Teams, phrases, patterns, and the language rule.
4. Warnings: length and `language.no_context`.

The send tools insert the final text from the review, never the raw input, then read the text back from the app and require it to match exactly before sending. The rules run again on the text that was read back.

### Language rule

The newest message from the other side is the context. It is cut at the first disclaimer or signature separator line, and URLs, email addresses, code spans, digits and @handles are removed before counting letters. When the context has at least 10 letters, its language is Arabic when at least half the letters are Arabic. The reply must be in the context language when that language is listed in `mirror`, otherwise in `default`. Without usable context, a reply that is not in `default` gets the warning `language.no_context`. `allow_language_mismatch=true` on a send tool skips the rule when the owner asked for another language.

## Error format

A refused send returns one message the model can fix in a single pass. Dashes in excerpts are shown as `[U+2014]` and similar, never as the character itself.

```
Not sent: the text breaks 2 persona rules. Fix them and call the tool again.
1. characters: forbidden character U+2014 EM DASH in body: "...the plan [U+2014] then...". Use a comma, period or colon.
2. signoff.trailing: nothing may follow "Best regards," (the signature card is attached automatically). Found "Alex".
```

## How the model sees it

- The server instructions end with a persona block when the `send` group is enabled: who the text is sent as, the hard rules in words for email and Teams, and the prose, capped at 4000 characters.
- `office_persona` returns the live state, version, prose, rules and rules in words.
- `office_check_text` runs the same review without sending and returns the final text, violations, fixes and warnings.
- Send tool descriptions end with the rules for their channel.

## Bootstrap from the owner's own writing

The `office persona-bootstrap` command reads the owner's own recent text and writes aggregate statistics to help draft a persona. It is read-only towards the apps: it reads Sent Items with unread state kept, and only Teams chats that are already read.

```
uv run --package office-mcp office persona-bootstrap --emails 40 --chats 15 --per-chat 20 --out ~/.config/office-mcp
```

- Output: `persona-stats.md`, mode 600, with the language mix and how often replies use the context language, median length per channel, top greetings and closings, dash and emoji counts, and at most 10 snippets with email addresses, phone numbers, URLs and capitalized name pairs replaced.
- The command refuses to write outside `~/.config/office-mcp/` and inside any git work tree.
- No corpus file is written unless `--keep-corpus` is given; that file, `persona-corpus.jsonl`, holds the raw text, is mode 600 and should be deleted after use.
- Draft `persona.md` from the statistics, review it, save it with `chmod 600`, delete `persona-stats.md`, and check that `office status` shows the persona as `loaded`.
