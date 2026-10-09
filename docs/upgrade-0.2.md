# Upgrading to Outlook CLI 0.2

Version 0.2 hardens existing OWA/REST workflows and adds opt-in local indexing. It does not require moving the default account to Graph or running a mailbox synchronization during installation.

## Changes scripts must account for

| Before | Version 0.2 |
|---|---|
| `-o file.json` saved raw data while stdout used an envelope | File and stdout both use the same envelope; use `--data-only` explicitly for old raw consumers |
| Some mutators and empty attachment lists emitted no JSON | All leaf commands support structured results and common output options |
| Parser errors could bypass JSON | Explicit JSON and piped parser errors emit `ok:false`, return nonzero |
| Search exhaustion looked complete | Unknown provider search scope is explicitly incomplete; a returned length is not a mailbox total |
| `summary.calendar.today_count` meant the displayed sample | `displayed_count` is the sample; totals are null unless completeness is known |
| Attachment JSON included base64 automatically | Metadata by default; `--include-content` opts into base64 |
| `read` accepted one ID | Multiple IDs return ordered per-item outcomes; one successful ID retains the email-object shape |
| JSON `id_map` writes could race, and old references were evicted | Transactional SQLite allocation, no 500-entry cap or number reuse |
| Authentication recovery could rerun an entire command | Only the rejected HTTP request may be retried; uncertain writes report `ambiguous_write` |
| Legacy scheduled items could be matched by subject | Cancellation requires a verified draft identity; unverified entries need inspection in Outlook |

The envelope `schema_version` remains `"1"`; use the CLI/package version to detect the 0.2 export behavior. Do not infer compatibility from schema_version alone.

```sh
# New default: file and stdout both contain {ok, schema_version, data, meta?}.
outlook search invoice --json -o results.json
jq '.data[] | .subject' results.json

# Explicit compatibility for an existing raw-array consumer.
outlook search invoice --data-only -o legacy-results.json
jq '.[] | .subject' legacy-results.json
```

Errors remain envelopes even with `--data-only`. Prefer the envelope for new scripts because raw exports omit completeness/error context on successful calls. Avoid `2>&1` when parsing stdout as JSON, and enable `set -o pipefail` when using Bash/Zsh pipelines.

## Installation and local state

Reinstall the package from the updated checkout, then inspect offline diagnostics:

```sh
pip install -e ".[dev]"
outlook doctor --json
outlook schema --json
```

`doctor` shows source and installed versions separately. It reads auth metadata without checking the keychain or making a network request. Normal mailbox access and optional `index sync` are separate actions.

Existing accounts, signatures and keychain tokens retain their profile locations. On first ID-store use, `id_map.json` is imported once into adjacent `id_map.sqlite3`; the original JSON is retained. Numbers remain account-scoped and are never reused or evicted. A number reflects its historical assignment, not the current list row. Always carry real `id` values from returned data in automation.

SQLite also stores optional local mail indexes in each selected profile's cache. Index creation is explicit; no daemon, timer or recurring sync is installed.

## Efficient read patterns

```sh
outlook inbox --limit 20 --view compact --no-input --json
outlook search 'from:alice@example.com' --all --fields id,subject,sender,received --json
outlook read REAL_ID_1 REAL_ID_2 --peek --body text --workers 4 --json
outlook attachments REAL_ID -d --save-to ./downloads --json
outlook draft-verify DRAFT_ID --json
```

`--limit` is a `--max/-n` alias. `--view compact` and selected fields are opt-in; existing full message output remains available. Mail list/search calls use `$select` for compact/selected fields. `--body-format` and `--body` control read-only mail representation, with `none|preview|text|html`; they do not replace compose BODY, `--body-file` or event body flags.

Use `--peek` to avoid `read` marking unread messages read. Bulk read/thread results keep input order and each item's success/error. Downloads now happen in JSON/piped mode and return file paths and statuses; existing files, symlinks and duplicate names are not overwritten.

## Completeness and failure semantics

Inspect `meta.complete`, `has_more`, `truncated_reason` and `returned_count`. List commands follow available continuation links, deduplicate IDs and bound runaway pagination. `--all` is available on inbox/search/folder/calendar/contacts/event-instances; it does not remove provider limits or safety budgets.

Provider search and subject-based thread lookup can return `complete:false`, `has_more:null`, `truncated_reason:"search_scope_unknown"` even when no nextLink exists. Do not use these results alone for authoritative counts. Summary unknown counts are null; failed sections return errors rather than zero.

Partial batch failures use `ok:false`, nonzero exit and completed/failed item details. Category propagation exposes resume checkpoints and stops on permanent failures or lack of progress. Retry an identical propagation only when its error says it is resumable.

`--no-input` prevents browser login and all interactive prompts. A confirmable mutation requires `-y`. `--dry-run` guards all mutation paths, including mark-read, event/category operations, account/signature changes and attachment downloads; read commands may still perform GETs. Recipients are normalized and validated before writes.

Transport refresh is request-scoped. A successful draft creation is not automatically repeated after a later 401. Safe reads have bounded transient retries; 429 honors `Retry-After`. An ambiguous write result must be inspected before retrying because the service may already have applied the action.

## Scheduled entries

New schedules store exact draft and tracking identities. Cancellation checks that the selected entry has not changed and that the identified message is still a draft, then deletes it before removing local tracking. Verification/network failures preserve tracking.

Old entries lacking an ID are shown as unverified and cannot be automatically cancelled. Inspect those entries in Outlook. Local tracking removal or matching subjects is no longer treated as proof that future delivery was stopped.

## Optional local index and Graph

```sh
outlook index sync --no-input --json
outlook index status --json
outlook index search invoice --domain example.com --limit 100 --json
outlook index search --folder Inbox --require-complete --json
```

Default REST sync scans full metadata/preview for every selected folder on each run; it is not an incremental Graph sync. Without `--folder`, all discovered folders and children are selected. `--include-body` explicitly stores bodies. Failed folders retain previous data and are marked incomplete. Search/status use no network and never refresh automatically.

Completeness refers to the explicitly indexed scope. Inspect folder timestamps, `oldest_sync`, `total_matches`, `returned_count`, `result_complete` and `body_scope`. `--require-complete` does not certify freshness or that all mailbox folders were chosen.

Graph is optional and currently limited to read/index operations. It requires an Entra public-client app with device login and delegated `Mail.Read`, `User.Read`, `offline_access` consent; tenant policy can require administrative approval.

```sh
outlook graph-login --client-id APP_UUID --tenant TENANT --account work
outlook index sync --backend graph --account work --json
outlook index search invoice --backend graph --account work --json
```

Use your own APP_UUID and TENANT. First Graph sync is full, later syncs use per-folder delta, and `--full` requests a rebuild. Invalid delta cursors are refreshed through a full folder sync. Default mail/calendar/mutation commands retain OWA/REST. Graph IDs are not interchangeable with REST IDs and should not be passed to those mutations.

## Diagnosis and measurements

`outlook schema COMMAND --json` exposes actual parameters without network access; quote multiword names such as `'index search'`. `outlook doctor --json` reports metadata and capabilities. `--profile` writes a separate stderr JSON record of execution/phase timings, request/retry counts and bytes without tokens or message bodies.

Measure representative tasks before claiming speedups. Compact output and metadata-only attachment fixtures can show output-size savings; they do not establish live mailbox latency or auth performance. Unit tests use offline fixtures and isolated state; smoke tests remain explicit live-account checks.
