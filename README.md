# outlook-cli

Read, search, send and manage Outlook mail, calendar, categories and attachments from the terminal. Version 0.2 adds consistent machine-readable output, safe attachment downloads, concurrent reads, persistent SQLite references and an optional local search index.

The `0.2.1.dev0` branch adds local-first mailbox research: compressed full text,
resumable copying, REST/Graph delta, Turkish-folded BM25 search and offline
conversation/attachment navigation. Existing command defaults remain compatible.

The default backend uses Outlook Web Access bearer tokens captured by Playwright, without configuring your own app registration. Optional Microsoft Graph indexing uses a separate app registration and consent flow. This is an unofficial project, unaffiliated with Microsoft. The existing REST v2 and internal OWA endpoints remain compatibility dependencies; see [LICENSE](LICENSE).

## Install and authenticate

Python 3.10+ is required. From this checkout:

```sh
pip install -e ".[dev]"
playwright install chromium
outlook login
outlook whoami --json
outlook doctor --json
```

For a published release, `pip install outlook365-cli` is also supported. Reinstall the editable package after changing versions so installed distribution metadata matches the source.

```sh
outlook account add work
outlook account list --json
outlook account switch work
outlook inbox --account work
outlook login --with-token < token.txt
```

Account selection is command `--account NAME`, `OUTLOOK_ACCOUNT`, persisted active account, then implicit `default`. A bound profile rejects authentication for a different mailbox. `--no-input` prevents prompts and browser login; it returns an authentication error when an existing session cannot be used. Interactive login remains an explicit option.

## Version 0.2 output contract

Read [the upgrade guide](docs/upgrade-0.2.md) before updating scripts.

Every leaf command accepts `--json`, `-o/--output`, and `--data-only`. Piped stdout automatically uses JSON. Human diagnostics and progress use stderr. Parser errors, runtime failures, empty results and mutating-command successes have structured output.

```json
{"ok":true,"schema_version":"1","data":[],"meta":{"returned_count":0}}
```

`meta` is included where applicable. Errors return `ok:false` and `error:{code,message}` with a nonzero exit status. Partial failures can include successful items in `data` as well as error details.

**File exports now match stdout.** `-o result.json` writes an envelope and stdout emits the same result. `--data-only` explicitly restores raw success data for legacy consumers; error output remains an envelope.

```sh
set -o pipefail
outlook search 'subject:report' --limit 20 --json -o results.json
outlook search 'subject:report' --limit 20 --data-only -o raw-results.json
outlook inbox --view compact --json |
  jq -e 'if .ok then .data else error(.error.message) end'
outlook folders --json | jq -e '.data[] | select(.name == "Inbox") | .unread_count'
```

Do not merge stderr into JSON. Preserve the CLI exit status in pipelines. For authoritative command options and offline installation diagnostics:

```sh
outlook schema --json
outlook schema read --json
outlook schema 'index search' --json
outlook doctor --json
outlook search invoice --profile --json
```

`doctor` reports source/distribution versions, account and stored auth metadata without opening the keychain or contacting Outlook. It does not prove the cached token is usable. `--profile` writes a separate stderr JSON record with elapsed time, phase times, request/retry counts, response bytes and status counts; it excludes message content and credentials.

## Read and search efficiently

```sh
outlook inbox --unread --limit 20 --view compact --json
outlook inbox --from alice --after 2026-09-01 --before 2026-10-01 --json
outlook inbox --has-attachments --no-category --body none --json
outlook folder Archive --all --fields id,subject,sender,received --json
outlook search 'from:alice@example.com' --all --view compact --json
outlook read REAL_ID_1 REAL_ID_2 --peek --workers 4 --body text --json
outlook thread REAL_ID_1 REAL_ID_2 --workers 4 --json
outlook summary --json
outlook open REAL_ID --print-url --json
```

`--max`, `--limit` and `-n` are aliases for count options. The default full message shape remains available. `--view compact` omits full bodies; `--fields` selects JSON fields. Message list/search commands request only the selected server fields where possible. Read-only mail commands accept `--body-format` / `--body` as `none`, `preview`, `text` or `html`. `html` preserves the provider's original body; `text` converts HTML to text. Compose BODY and calendar `--body` keep their existing meanings.

`read` normally marks unread messages read. Use **`--peek` for side-effect-free mail research**. `--dry-run read` also disables auto-marking. A successful single read retains the email-object shape; a multi-ID read returns ordered items with `id`, `ok`, `data` and optional `error`. Thread reads follow the same single/multiple distinction. Both support 1–4 workers. Batch failures preserve successful items and return nonzero status.

### Completeness and counts

`--all` follows available pages for inbox, folder, search, calendar, contacts and event-instances, subject to bounded page/result safeguards. Pagination deduplicates provider IDs and reports `complete`, `has_more`, `truncated_reason`, `pages`, `fetched_count`, `returned_count`, `deduplicated_count` and `filtered_count`.

Provider search can stop returning continuation links without proving whole-mailbox coverage. Such results report `complete:false`, `has_more:null` and `truncated_reason:"search_scope_unknown"`. Thread lookup searches by subject and filters conversation IDs, so it also cannot guarantee a globally complete conversation. An exhausted provider search is not a certified total.

`summary` distinguishes displayed messages/events from known totals. Unknown totals are null, and failed sections are explicitly reported with a nonzero exit status. It never turns a failed API request into a successful zero count.

### Stable references

Use the returned real `id` in scripts. `display_num` is a persistent account-local reference, not the position in the latest result list. Five returned messages need not have numbers 1–5.

IDs now live in transactional SQLite with concurrent allocation, legacy JSON migration and no 500-entry eviction or number reuse. Each account has its own reference space. REST and Graph IDs are separate provider namespaces.

## Attachments and drafts

```sh
outlook attachments REAL_ID --json
outlook attachments REAL_ID -d --save-to ./downloads --json -o downloads.json
outlook attachments REAL_ID --include-content --json
outlook draft-verify DRAFT_ID --json
```

Attachment listing fetches metadata without base64 content. `-d` downloads before formatting output, so it works in pipes and JSON mode. Each result reports `status`, and successful downloads include absolute `path` and `bytes_written`. Unsafe filename path components are removed; existing files, symlinks and duplicate filenames are refused instead of overwritten. Failures remain visible per attachment. `--include-content` deliberately opts into potentially large base64 JSON output.

`draft-verify` reads actual To/CC recipients and attachment metadata in one command without marking the message read.

## Compose, send and manage

```sh
outlook draft 'a@example.com;b@example.com' Subject --body-file message.txt \
  --cc 'c@example.com,d@example.com' --cc e@example.com --json
outlook send a@example.com Subject Body --to b@example.com --dry-run --json
outlook send a@example.com Subject --body-file message.txt -a report.pdf -y --no-input --json
outlook reply REAL_ID 'Thanks!' --all --dry-run --json
outlook reply-draft REAL_ID 'Will review tomorrow' --json
outlook draft-send DRAFT_ID -y --no-input --json
outlook forward REAL_ID colleague@example.com --comment FYI --dry-run --json
outlook categorize ID_1 ID_2 Finance --json
outlook uncategorize ID_1 ID_2 Finance --json
outlook move ID_1 ID_2 Archive --dry-run --json
outlook copy ID_1 ID_2 Archive --dry-run --json
outlook mark-read ID_1 ID_2 --unread --json
outlook delete ID_1 ID_2 -y --no-input --json
outlook flag REAL_ID --due tomorrow --json
outlook pin REAL_ID --unpin --json
```

TO/CC parsing accepts commas, semicolons and repeated options, removes duplicate addresses, and validates locally before API writes. `--to` adds to the required positional TO. `--attach/-a` is repeatable. `--body-file -` reads the body from stdin once; `--html` handles explicitly supplied HTML. Plain-text multiline bodies preserve line breaks when converted for Outlook. `--signature NAME` appends a saved signature.

Draft commands do not send mail. Send/delete and other confirmable operations require confirmation or `-y`; in `--no-input` mode they require `-y`. `--dry-run` previews mutating commands before executing changes, including local account/signature/index operations. Read-only commands may still fetch data. `--enable-commands inbox,read,search` restricts enabled top-level commands.

Management/category batches report individual successes and failures. Category propagation checkpoints incomplete operations and stops after bounded retries or lack of progress. When an error explicitly identifies a resumable checkpoint, repeating the same command resumes it.

## Scheduling

```sh
outlook schedule a@example.com Subject Body '+1h' --dry-run --json
outlook schedule-draft DRAFT_ID 'tomorrow 09:00' -y --json
outlook schedule-list --json
outlook schedule-cancel 1 -y --json
```

Times accept `+30m`, `+1h`, `+2h30m`, `today 17:00`, `tomorrow 09:00` and ISO dates/times. New schedules persist a draft identity and tracking identity before sending. Cancellation rechecks the selected tracking identity and verifies the server item is still a draft before deletion. It preserves tracking after failures.

Legacy entries without a verified message ID are marked unverified/uncancellable. The CLI no longer guesses a draft by subject or treats deleting local tracking as proof that sending was cancelled. Check those entries in Outlook.

## Calendar, contacts and categories

```sh
outlook calendar --days 7 --all --timezone Europe/Istanbul --json
outlook calendar --days -7 --calendar 'Team' --json
outlook calendars --json
outlook event EVENT_ID --json
outlook event-create Meeting 'tomorrow 14:00' 'tomorrow 15:00' \
  -a colleague@example.com --teams --dry-run --json
outlook event-create Standup 'tomorrow 09:00' 'tomorrow 09:15' \
  --repeat weekly --repeat-count 8 --dry-run --json
outlook event-update EVENT_ID --subject 'Updated title' --dry-run --json
outlook event-delete EVENT_ID --series --dry-run --json
outlook event-respond EVENT_ID accept --dry-run --json
outlook event-instances EVENT_ID --days 90 --all --json
outlook free-busy colleague@example.com tomorrow -d 30 --json
outlook people-search Alice --limit 5 --json
outlook contacts --limit 100 --all --json
outlook categories --json
outlook category-create Finance --color 7 --dry-run --json
outlook category-rename Old New --dry-run --json
outlook category-clear Finance --folder Inbox --dry-run --json
outlook category-delete Finance --dry-run --json
outlook signature-list --json
outlook signature-show default --json
outlook signature-pull --name work --dry-run --json
```

Calendar day windows use local midnight boundaries; `--timezone` controls serialized calendar times and accepts IANA names or fixed UTC offsets. Calendar creation, updates, recurring instances, meeting responses and shared-calendar lookup continue using the default REST backend.

## Local research index

For the new full-mailbox workflow, use `local`:

```sh
outlook local import-index --no-input --json  # optional; preserves index.sqlite3
outlook local sync --no-input --json          # text + attachment metadata, then delta
outlook local status --json
outlook local search 'toplantı domain:example.com after:2026-09-01' --json
outlook local search 'çalışmalarından' --match stem --has-attachments --json
outlook local read MESSAGE_ID --json           # offline text; never marks read
outlook local thread MESSAGE_OR_CONVERSATION_ID --json
outlook local related MESSAGE_ID --json
outlook local attachments MESSAGE_ID --json
```

Search returns bounded snippets, `[[highlighted matches]]`, BM25 scores and IDs
without full bodies. Research reads use no network/keychain or automatic refresh.
Sync selects every discovered primary-mailbox folder and child by default; it
does not certify a separate online archive or recovery store. Attachment binaries
are downloaded separately with the existing attachment command.

The new store is profile-scoped `mail.sqlite3` under `~/.cache/outlook-cli/`.
Database/WAL/SHM are 0600; the directory is 0700. It is not application-encrypted.
`local purge -y --no-input` removes the new store and its sidecars, preserving
`index.sqlite3`. `local compact` reclaims superseded bodies and free pages.
See [the local guide](docs/local-mail.md) and [design research](docs/local-store-research.md).

The legacy `index` commands retain their behavior:

```sh
outlook index sync --account work --no-input --json
outlook index sync --folder Inbox --folder Archive --no-input --json
outlook index status --account work --json
outlook index search invoice --domain example.com --limit 100 --json
outlook index search '"project alpha"' --from alice@example.com --after 2026-09-01 --json
outlook index search --folder Inbox --require-complete --json
```

Sync defaults to every discovered folder and child folder. Explicit `--folder NAME_OR_ID` restricts scope. A REST sync takes a complete metadata/preview snapshot of each selected folder on every run; it does **not** use delta. `--include-body` opts into storing bodies for full-text search. Successful folder snapshots replace old membership, so deleted/moved messages are reconciled; failed folders retain previous data and become incomplete.

Local search/status do not contact Outlook or automatically sync. Queries support exact sender/recipient/domain, dates, conversation IDs and quoted phrases. Results report folder scope/timestamps, `oldest_sync`, `total_matches`, `returned_count`, `has_more` and `result_complete`. `--require-complete` checks the explicitly indexed scope, not whether the data is current or every mailbox folder was selected. No scheduler or background daemon is installed.

### Optional Graph delta

Graph indexing requires a separately configured Entra public-client application that permits device login and delegated `Mail.Read`, `User.Read` and `offline_access` consent. Tenant policy may require an administrator. Existing OWA authentication is not automatically converted into Graph authentication.

```sh
outlook graph-login --client-id APP_UUID --tenant TENANT --account work
outlook index sync --backend graph --account work --json
outlook index search invoice --backend graph --account work --json
outlook index sync --backend graph --full --account work --json
```

Replace APP_UUID and TENANT with your own application's values. `OUTLOOK_GRAPH_CLIENT_ID` and `OUTLOOK_GRAPH_TENANT` can supply defaults. Device login is explicit and cannot run with `--no-input`; existing Graph refresh credentials can be reused without interactive login.

The first Graph sync is full; subsequent syncs use per-folder delta checkpoints committed with data. Invalid delta cursors trigger a full folder refresh. `--full` forces a rebuild. Graph read requests use immutable IDs. **Graph currently supports read/index operations only**; send, calendar, schedule, pin and category mutations keep their current REST/OWA paths. Do not pass Graph index IDs to those REST commands.

## Authentication, transport and storage

`browser.headless: true` opts into automatic browser renewal with saved SSO state
and a maximum 30-second capture window. Explicit `login`/`account add` remain
visible flows. `--no-input` never opens a browser, including a headless one.

401 recovery retries only the rejected HTTP request. The CLI never reruns an entire command after authentication failure. Safe reads retry selected transient/network failures within a bounded retry budget; explicit 429 respects `Retry-After`. Writes with uncertain network/5xx outcomes return `ambiguous_write` rather than being replayed. Inspect the affected item before retrying. Per-account locks limit concurrent requests across processes and serialize refresh/state updates.

Cache defaults to `~/.cache/outlook-cli/`, config to `~/.config/outlook-cli/`; override with `OUTLOOK_CLI_CACHE` and `OUTLOOK_CLI_CONFIG`. Named profiles use `accounts/<profile>/` below each root. Legacy implicit `default` paths remain supported. Key files include `id_map.sqlite3`, `scheduled.json`, `index.sqlite3`, `browser-state.json` and non-secret `token.json` metadata. Existing `id_map.json` is migrated once and retained as legacy input.

Bearer tokens live in the OS keychain. Browser session state, local email indexes and signature files are private user data and should not be committed or shared. Graph credentials use their own keychain service. `account remove NAME` manages profile removal; removing local cache files alone does not revoke server-issued tokens.

Global `config.yaml` and account-level overrides support:

```yaml
max_messages: 25
default_signature: null
timezone: UTC
browser:
  headless: false
  timeout: 120
```

## Exit codes and development

Codes: `0` success; `1` failure or partial batch; `2` invalid usage; `4` authentication required; `5` not found; `7` rate limited; `8` retryable transport; `10` account/config; `130` interrupted. A partial batch uses its item errors for detail rather than mapping the entire batch to one item's code.

```sh
pytest -m 'not smoke'
pytest -m smoke  # explicitly uses a live account
```

Unit tests isolate cache/config and deny live network, keychain and browser access unless mocked. New transport, pagination, concurrent ID allocation, output contracts and index tests use offline fixtures. See [AGENTS.md](AGENTS.md) for implementation boundaries and [SKILL.md](SKILL.md) for concise AI Agent usage.
