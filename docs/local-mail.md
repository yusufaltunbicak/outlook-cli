# Local mail research

## Setup and research

```sh
outlook local import-index --no-input --json  # optional sideways import
outlook local sync --no-input --json
outlook local search 'toplantı' --limit 10 --json
outlook local search 'calis' --folder Inbox --has-attachments --json
outlook local search 'çalışmalarından' --match stem --json
outlook local search '"proje alpha" person:alice@example.com after:2026-09-01' --json
outlook local read MESSAGE_ID --json
outlook local thread MESSAGE_OR_CONVERSATION_ID --limit 100 --json
outlook local related MESSAGE_ID --json
outlook local attachments MESSAGE_ID --json
```

Installation never copies mail automatically. Sync intentionally copies every
discovered primary-mailbox folder and child, including sent, deleted and draft
items. Repeated `--folder NAME_OR_ID` narrows scope. Discovery refreshes every
sync, handling new/renamed/deleted folders. It does not certify separate online
archive or recoverable-item mailboxes. Empty discovered folders count as synced.

Stored fields include text bodies, sender/To/CC/BCC/reply-to, received/sent/modified
dates, conversation and InternetMessageId, and attachment metadata. New REST copies
request text at the server. Legacy-imported HTML can be retrieved with
`local read ID --html`. Attachment bytes are not copied; names are searchable,
file contents are not. Use existing `outlook attachments ID -d --save-to DIR`
for relevant REST attachment files. Text is not an exact MIME/HTML export.

Search/read/thread/related/attachment metadata are offline with no auth/keychain,
network or automatic refresh. `local read` never marks mail read. The older
network `read` still needs `--peek` for side-effect-free research. Sync explicitly
when freshness matters, then research repeatedly without another network call.

## Query language and output

Unquoted words are ANDed; quotes mean adjacent-word phrases. Fields: `from`, `to`,
`person`, `domain`, `folder`, `after`, `before`, `has:attachment(s)`, and
`thread`/`conversation`. Quote values containing spaces (`folder:"Sent Items"`).
Explicit flags override corresponding fields. Addresses/domains match exactly;
person also matches folded display-name fragments. `after` is inclusive and
`before` exclusive; bare dates mean UTC midnight. Boolean operators/regex are
not interpreted.

Normalization folds ı/i, İ/I, ş/s, ğ/g, ü/u, ö/o, ç/c and case, preserving original
display spelling. Default prefix search supports partial word starts. Exact mode
selects whole tokens; stem mode strips a small Turkish suffix set before prefix
search. This can broaden results and does not promise every root or consonant
alternation. Quoted phrases do not expand. Arbitrary infix matching and full
linguistic stemming are future options.

FTS5 uses BM25 with higher weights for subjects/people than body. Search returns
short snippets and `[[...]]` highlights without full bodies. Trim with
`--fields id,subject,snippet,highlighted`. Fetch a full message/thread only for
relevant results. Thread order is chronological; metadata reports truncation and
stored-scope completeness. Defaults of existing commands remain unchanged.
`--view compact` preserves local snippets and highlights. Turkish apostrophes
within words, such as `İstanbul'da`, are accepted. Attachment-list metadata
reports `metadata_complete`; legacy imports remain unverified until a full sync.

## Resume, API limits and coverage

REST uses native `Prefer: odata.track-changes`; its initial deltaLink is a seed
that must be followed before copying finishes. Later runs use final per-folder
cursors. Graph uses `/messages/delta`, immutable IDs and separate authentication.
Changed records resolve against current server state; tombstones remove only
originating folder membership. Cursors are never emitted. Repeated entries do
not duplicate local messages; identical bodies share compressed storage while
distinct mailbox messages remain distinct even with the same InternetMessageId.

Each page and continuation commit together. Repeat a command to resume;
`--full` restarts replacement snapshots. Existing complete generations remain
searchable until replacement completes. Invalid delta state gets one full folder
refresh. Servers without tracking use honest full snapshots labelled `snapshot`.

Sync is serial, at least 0.25 seconds between requests, normally 100 messages per
page. Shared HTTP transport observes `Retry-After`, bounded safe-read retries and
per-account concurrency locks. Sync stops at failure rather than hitting more
folders. Long 429 cooldowns persist and are checked before another request.
Ctrl-C exits 130; committed pages survive. `--max-pages` bounds work without
discarding continuation state.

`local status` distinguishes `whole_mailbox_complete` (current discovered folder
inventory) from selected-scope `complete`. `oldest_sync` describes age; it cannot
prove current remote freshness. New mail/moves can occur immediately afterward.
A failed delta may expose committed changed pages with incomplete scope metadata.

## Storage and deletion

Default: `~/.cache/outlook-cli/mail.sqlite3`; named profiles use
`~/.cache/outlook-cli/accounts/PROFILE/mail.sqlite3`. `OUTLOOK_CLI_CACHE` applies.
Database/WAL/SHM permissions are 0600; directory is 0700. Tokens stay in the OS
keychain. Text exists in compressed bodies and searchable FTS structures.
Permissions/compression are not encryption; no application encryption or launchd
service is installed.

```sh
outlook local compact --no-input --json      # reclaim old bodies/free pages
outlook local purge --dry-run --json         # preview
outlook local purge -y --no-input --json     # remove new store and sidecars
outlook local purge --include-legacy-index -y --no-input --json  # explicit: both mail copies
```

Default purge preserves legacy `index.sqlite3`, credentials, signatures and ID
maps. The legacy index also contains email. Only the explicit
`--include-legacy-index` flag authorizes removing both stores and their sidecars
in one command; use it after reviewing the migration. Without `-y`, confirmation
is required; no-input mode refuses without `-y`. File deletion is not
cryptographic SSD/backup erasure.

Import opens the legacy source read-only, retaining source coverage/freshness;
missing bodies/attachment metadata are not invented. A following full sync
replaces imported generations. Compact afterward to reclaim superseded HTML.

## Measurements and remaining ideas

`scripts/benchmark_local.py` reports content-free old/new query and CLI timings,
sizes and saved real sync aggregates. CLI process startup is separate from warm
in-process search. Stored scopes differ and are reported. Logical successful GET
counts exclude transport retries; Graph wire bytes are unavailable.
`scripts/benchmark_search_engines.py` reproduces synthetic word/trigram comparison.
Measured results: [local-mail-benchmark.md](local-mail-benchmark.md).
Sources and tradeoffs: [local-store-research.md](local-store-research.md).

Next priorities are supported Graph deployment when app setup is justified,
corpus-based Turkish morphology/infix recall evaluation, opt-in idle background
sync if stale data becomes inconvenient, attachment text extraction, then semantic
reranking after relevance evaluation. The current workflow needs no model download
or daemon. Explicit sync keeps offline querying predictable.
