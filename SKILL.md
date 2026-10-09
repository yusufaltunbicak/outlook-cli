---
name: outlook-cli
description: Read, search and manage Outlook mail, calendar, categories and attachments with the local outlook CLI. Supports account-scoped local search and optional Graph indexing.
author: yusufaltunbicak
version: "0.2.1.dev0"
tags: [outlook, email, office365, calendar, attachments, cli]
---

# Outlook CLI for AI Agents

Use the installed `outlook` command. Discover exact options with `outlook schema COMMAND --json` or `outlook COMMAND --help`; do not guess flags. The source checkout is the authority for this skill. See `docs/upgrade-0.2.md` for compatibility changes.

## First checks

```bash
outlook doctor --json
outlook schema read --json
outlook schema 'index search' --json
outlook account list --json
outlook whoami --account work --no-input --json
```

`doctor` reads metadata, not the keychain, and does not verify a live token. Named account selection: command `--account NAME`, `OUTLOOK_ACCOUNT`, persisted active profile, then `default`. `schema` is account-independent.

## Output contract

Piped stdout is automatically JSON. `--json` makes this explicit. Status/progress messages and `--profile` measurements go to stderr. Never merge stderr into JSON or discard it to conceal errors.

```bash
set -o pipefail
outlook inbox --max 10 --view compact --no-input --json |
  jq -e 'if .ok then .data else error(.error.message) end'
outlook folders --json | jq -e '.data[] | select(.name == "Inbox") | .unread_count'
outlook search 'subject:report' --limit 20 --json -o results.json
```

Success: `{ "ok": true, "schema_version": "1", "data": ..., "meta": ... }`. Errors have `ok:false`, `error.code`, `error.message`, and a nonzero exit code; partial failures can also contain successful items in `data`. `meta` is present where applicable. Empty lists produce `data:[]`.

**Version 0.2:** `-o/--output` writes the same envelope as stdout. Use `--data-only` explicitly for legacy raw success data; errors remain structured. Check both exit status and `ok`. Prefer envelopes when completeness matters.

## Efficient mail research

```bash
outlook inbox --unread --limit 20 --view compact --json
outlook inbox --from alice --after 2026-09-01 --body none --json
outlook search 'from:alice@example.com' --all --fields id,subject,sender,received --json
outlook read REAL_ID_1 REAL_ID_2 --peek --workers 4 --body text --json
outlook thread REAL_ID_1 REAL_ID_2 --workers 4 --json
outlook summary --json
```

Use real `id` values from results. `display_num` is a persistent account-local reference, **not the result row number**; listing five messages does not imply IDs 1–5. SQLite allocation is concurrent-safe, has no 500-entry eviction, and does not reuse numbers. Graph and REST provider IDs are not interchangeable.

`--max`, `--limit`, `-n` are aliases where a count option exists. `--view compact` and `--fields` reduce message output; list/search requests also use server-side field selection. Read-only mail commands accept `--body-format` / `--body` with `none|preview|text|html`; compose commands retain their existing BODY / `--body-file` semantics.

`read` normally marks unread mail as read. Use **`--peek` for research**. `--dry-run read ...` also suppresses auto-marking. One successful `read` returns an email; multi-ID reads return ordered `{id,ok,data,error?}` items. A failed item does not erase earlier successes.

Pagination metadata includes `complete`, `has_more`, `truncated_reason`, `pages`, `fetched_count`, `returned_count`. `--all` follows available pages on inbox/search/folder/calendar/contacts/event-instances, subject to safety limits. Provider search and subject-based thread lookup report `complete:false` / `search_scope_unknown` when global completeness cannot be proven. Never report the returned count as a whole-mailbox total. `summary` distinguishes `displayed_count` from total counts and uses null for unknown totals.

## Local index for repeated research

Prefer `local` for repeated research after the user authorizes initial storage:

```bash
outlook local status --json
outlook local search 'toplantı domain:example.com after:2026-09-01' --limit 10 --json
outlook local search 'çalışmalarından' --match stem --has-attachments --json
outlook local search '"proje alpha" person:alice@example.com' --json
outlook local read MESSAGE_ID --json
outlook local thread MESSAGE_OR_CONVERSATION_ID --limit 50 --json
outlook local related MESSAGE_ID --limit 10 --json
outlook local attachments MESSAGE_ID --json
```

These reads are offline and never mark mail read: no auth/keychain or network.
Carry real IDs and backend; REST/Graph IDs have different namespaces. Search is
already compact with bounded snippets, BM25 score and `highlighted` using `[[...]]`.
Use `--fields id,subject,snippet,highlighted` to trim further. Avoid `--view compact`
on local snippets because the legacy compact projection removes those fields.
Fetch full messages/threads only when a result is relevant.

Literal words are ANDed; quotes mean exact phrases. Fields are `from:`, `to:`,
`person:`, `domain:`, `folder:`, `after:`, `before:`, `has:attachments`, `thread:`
or `conversation:`. Quote field values containing spaces (`folder:"Sent Items"`).
Explicit flags are also supported. Turkish/case folding is automatic. Prefix is
default; `--match exact` selects whole words and `--match stem` adds conservative
suffix stripping. Stem is a heuristic, without full morphology/infix guarantees.

Check `whole_mailbox_complete`, `complete`, `oldest_sync`, `has_more` and
`result_complete`: coverage is as of sync, not current remote freshness. No
automatic refresh occurs; a separate online archive and attachment bytes are
outside the default copy. A local thread can be partial when folders are missing.

```bash
outlook local import-index --no-input --json  # optional; source retained
outlook local sync --no-input --json          # all discovered folders; text+metadata
outlook local sync --folder Inbox --no-input --json
outlook local sync --backend graph --no-input --json  # separately configured Graph
```

Sync checkpoints each page. Repeat a failed/interrupted command to resume;
`--full` deliberately restarts replacement snapshots. REST delta is a deprecated
API compatibility path; Graph needs explicit registration/login. Sync is serial,
paced, stops on failure and honors 429/Retry-After, persisting long cooldowns.
Never log tokens, cursors or unnecessary message content.

Data is in profile cache `mail.sqlite3` (0600 DB/WAL/SHM, 0700 directory). It is
not application-encrypted. `local purge -y --no-input` removes only the new store;
the legacy index survives. Delete only when authorized. No daemon is installed.
`local purge --include-legacy-index -y --no-input` explicitly deletes both local
mail stores and sidecars. Never add that flag without the user's authorization.

Legacy `index` commands below retain their 0.2 behavior:

```bash
outlook index sync --account work --no-input --json
outlook index status --account work --json
outlook index search invoice --account work --domain example.com --limit 100 --json
outlook index search '"project alpha"' --from alice@example.com --after 2026-09-01 --json
outlook index search --folder Inbox --require-complete --json
```

Sync defaults to all discovered folders and children, including available sent/archive folders. `--folder NAME_OR_ID` can be repeated. REST sync re-fetches full folder metadata/preview each time; it is not delta. Local search/status use no network or auth and never auto-refresh. Inspect folder timestamps, scope, `total_matches`, `returned_count`, `result_complete`; `--require-complete` checks synced scope, not freshness or whole-mailbox coverage. Full body search requires an explicit earlier `index sync --include-body`.

Graph delta indexing is optional: `graph-login --client-id APP_UUID --tenant TENANT` needs an Entra public-client app and delegated `Mail.Read`, `User.Read`, `offline_access` consent, subject to tenant policy. Then use `index sync --backend graph` and `index search --backend graph`. First sync is full, later ones use per-folder delta; `--full` rebuilds. Graph is currently read/index-only. Default mail, sending, calendar and OWA operations continue using existing authentication. Do not feed Graph index IDs into REST mutation commands.

## Attachments

```bash
outlook attachments REAL_ID --json
outlook attachments REAL_ID -d --save-to ./downloads --json -o downloads.json
outlook draft-verify DRAFT_ID --json
```

Listing fetches metadata by default. Downloads work with pipes and JSON; results include per-file `status`, `path`, `bytes_written` or `error`. Unsafe path components are removed and existing files, symlinks and duplicate filenames are not overwritten. Use another download directory after a conflict. `--include-content` explicitly requests base64 in JSON and can produce very large output. `draft-verify` reads actual recipients and attachment metadata without marking mail read.

## Compose and change mail

Only send or otherwise communicate when the user's task authorizes it. Creation of an unsent draft does not send mail.

```bash
outlook draft 'a@example.com;b@example.com' 'Subject' --body-file message.txt \
  --cc 'c@example.com,d@example.com' --cc e@example.com --json
outlook send a@example.com 'Subject' 'Body' --to b@example.com --dry-run --json
outlook send a@example.com 'Subject' --body-file message.txt -a report.pdf -y --no-input --json
outlook reply REAL_ID 'Thanks' --all --dry-run --json
outlook reply-draft REAL_ID 'Will review' --json
outlook draft-send DRAFT_ID -y --no-input --json
outlook forward REAL_ID colleague@example.com --comment FYI --dry-run --json
outlook categorize ID_1 ID_2 Finance --json
outlook move ID_1 ID_2 Archive --dry-run --json
outlook delete ID_1 ID_2 -y --no-input --json
```

TO/CC accept commas, semicolons and repeated `--to` / `--cc`; invalid addresses fail before API calls. `--to` adds recipients to the required positional TO. `--attach/-a` is repeatable. `--body-file -` reads stdin once. `--dry-run` previews mutations before changing remote or local state; read-only requests can still occur for read commands. `--no-input` disables prompts and browser login; confirmable actions require `-y`. `--enable-commands` limits top-level commands.

## Scheduling, calendar, categories

```bash
outlook schedule a@example.com 'Subject' 'Body' '+1h' --dry-run --json
outlook schedule-draft DRAFT_ID 'tomorrow 09:00' -y --json
outlook schedule-list --json
outlook schedule-cancel 1 -y --json
outlook calendar --days 7 --all --timezone Europe/Istanbul --json
outlook event EVENT_ID --json
outlook event-create Meeting 'tomorrow 14:00' 'tomorrow 15:00' -a a@example.com --dry-run --json
outlook free-busy a@example.com tomorrow -d 30 --json
outlook categories --json
outlook category-rename Old New --dry-run --json
outlook signature-list --json
```

Schedule cancellation verifies the stored draft identity; legacy entries without a verified ID cannot be cancelled automatically or by subject guessing. Check those in Outlook. Category propagation records bounded failures and a resume checkpoint; only repeat the identical command when its error explicitly says it is resumable.

## Authentication and failures

`outlook login` is interactive OWA login; `outlook login --with-token < token.txt` is explicit token input. Tokens are stored in the OS keychain, and `token.json` contains metadata. Account-scoped browser state, IDs, schedules, signatures and indexes remain local; never expose tokens or browser-state contents.

401 refresh retries only the rejected HTTP request; it never replays the entire command. Safe reads have bounded transient retries; 429 respects `Retry-After`. An `ambiguous_write` must be inspected before retrying because the server may already have applied it. `--no-input` returns an auth error when login is needed instead of opening a browser.

Exit codes: 0 success, 1 failure/partial, 2 invalid usage, 4 auth, 5 not found, 7 throttled, 8 retryable transport, 10 account/config, 130 interrupted. Bulk commands use nonzero failure status with per-item details. Use `outlook COMMAND --profile` for stderr timing, request/retry counts and bytes; do not claim a speedup without measurement.
