# Local mailbox design research

Reviewed against primary sources on **2026-10-09**. This is a design record;
synthetic experiments below are not measurements of a production mailbox.

## Decision

Use the existing OWA session with Outlook REST v2 **native per-folder change
tracking**, a separate account-local SQLite store, transactional page checkpoints,
compressed body blobs, contentless word-token FTS5, and explicit synchronization.
Keep the existing Graph
backend available for accounts that already have an app registration. Do not
create a registration or install a background service implicitly.

This minimizes setup for the existing account. It does not make REST v2 a supported
public API: Microsoft marks it deprecated and its archived documentation announced
decommissioning in March 2024. Continued success with this account's OWA token is
an observed compatibility path, not a service guarantee. A supported long-term
backend is Microsoft Graph. [Microsoft REST status](https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api)

## Backend comparison

| Path | Authentication / permissions | Changes and deletion tracking | Limits / fragility | Choice |
|---|---|---|---|---|
| REST v2 using current OWA token | Existing browser session; legacy documented read scope is `mail.read` | `Prefer: odata.track-changes` on folder messages | Deprecated; verify capability for each folder; no assumption of a current supported quota | Default compatibility path |
| OWA internal `service.svc` | First-party browser token/session; no supported third-party permission contract | Undocumented action/schema behavior | Internal implementation can change; no public supported quota | Do not add another reverse-engineered sync dependency |
| Graph v1.0 | Own Entra public client; delegated `Mail.Read` for bodies/attachments, `offline_access` for refresh; existing CLI also uses `User.Read` | Supported folder message delta and folder hierarchy delta | Documented per mailbox/app quotas; consent can depend on tenant policy | Supported optional path |
| IMAP OAuth | Entra registration + `IMAP.AccessAsUser.All`, usually `offline_access`; mailbox protocol must be enabled | UID-based enumeration; optional server extensions require capability checks | Another authentication/sync stack; less Exchange-specific thread metadata | Possible export/interoperability backend, not initial implementation |
| EWS | OAuth registration and EWS access permitted by tenant | Mature `SyncFolderItems` protocol | Exchange Online retirement has begun; tenant-dependent availability | Reject for new work |

Graph message delta can use `Mail.ReadBasic` for limited metadata; full bodies and
attachments need `Mail.Read`. Changes include folder departures, so a tombstone
must remove membership in the originating folder, rather than blindly deleting a
message already seen in another folder. Persist opaque continuation links as
returned, and reapply request headers. [Graph message delta](https://learn.microsoft.com/en-us/graph/api/message-delta?view=graph-rest-1.0),
[Graph permissions](https://learn.microsoft.com/en-us/graph/permissions-reference#mailread)

Graph immutable IDs require `Prefer: IdType="ImmutableId"` on every request. They
survive moves within one mailbox, but not moves to a separate archive mailbox or
export/reimport. Graph and REST IDs must remain backend-scoped.
[Immutable IDs](https://learn.microsoft.com/en-us/graph/outlook-immutable-id)

Graph Outlook quotas currently list **10,000 requests / 10 minutes** and **four
concurrent requests** per app/mailbox. The 150 MB / five-minute quota is for
uploads, not download allowance. These are ceilings, not throughput targets.
Use a serial, paced first sync and bounded retries; honor `Retry-After` on 429 and
back off for transient read failures. A JSON batch does not evade limits.
[Service quotas](https://learn.microsoft.com/en-us/graph/throttling-limits#outlook-service-limits),
[Retry guidance](https://learn.microsoft.com/en-us/graph/throttling)

Graph delta state can expire without a fixed Outlook token lifetime; 410 or
`syncStateNotFound` requires a fresh folder snapshot. Keep old searchable data
until replacement completes, and treat repeated changes idempotently.
[Delta lifecycle](https://learn.microsoft.com/en-us/graph/delta-query-overview)

IMAP requires its own resource-scoped OAuth token and XOAUTH2; the existing OWA
token is not assumed usable. A strictly pull-only client would use read-only
mailbox selection and `BODY.PEEK`, never flag/store/expunge commands.
[IMAP OAuth](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth)

### EWS status at the review date

The original 2023 announcement said blocking would start on 2026-10-01. The
Exchange team's **2026-02-05** update changed this into phased, admin-controlled
disablement starting October 2026 and concluding in a full shutdown in **2027**.
Default/unconfigured tenants can be switched to disabled as rollout proceeds;
explicit enabled tenants need the app allow list for continued access. The allow
list announcement was updated **2026-09-22**. Therefore, on October 9 it is wrong
to assert that every tenant has already lost EWS, or that EWS is a durable new
dependency. Exchange Server on premises is outside this retirement.
[Updated retirement plan](https://techcommunity.microsoft.com/blog/exchange/exchange-online-ews-your-time-is-almost-up/4492361),
[Allow-list changes](https://techcommunity.microsoft.com/blog/exchange/introducing-ewsallowedappids-preparing-for-the-final-phase-of-ews-retirement/4529471)

### REST implementation details that matter

The legacy endpoint is folder `/messages` with `Prefer: odata.track-changes`,
**not** Graph's `/messages/delta`. Confirm `Preference-Applied`. A legacy initial
response's delta link is a seed: follow it once to continue initial enumeration,
then follow next links until the final delta link. `$select`, `$top`, and `$expand`
are documented; ordinary free-text `$search` is not a delta option.

Legacy tombstones can be `{ "id": "Messages('ID')", "reason": "deleted" }`.
Support that lowercase resource-wrapped ID as well as modern `@removed` fixtures.
Metadata-only attachment listing uses `$select` to exclude `ContentBytes`.
Validate continuation URL host and API path before attaching a bearer token.
[Legacy sync and attachments](https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations)

Sparse changed records must preserve existing properties, or be hydrated by a GET
before serialization. Microsoft's sample explicitly merges changed fields.
Never convert an absent `Body` to an empty body and overwrite the local copy.
[Microsoft synchronization lab](https://github.com/OfficeDev/hands-on-labs/blob/master/O3653/O3653-22%20Synchronize%20Outlook%20data/Lab.md)

If change tracking is absent, a full folder snapshot is an honest fallback.
A `LastModifiedDateTime` watermark alone cannot detect deletions, moves out, or
all hierarchy changes. Never label such polling a complete delta backend.
An EWS operation named `SyncFolderItems` does not establish that an OWA JSON
service supports that operation with a stable schema.

## Prior art and what to borrow

| Project | Relevant design | Borrow / avoid |
|---|---|---|
| [wacli](https://github.com/openclaw/wacli), created by steipete | Local SQLite app store plus separate protocol session; offline search; FTS5; store locking and coverage reporting | Separate authentication from searchable data; explicit sync and offline research; report coverage honestly |
| [notmuch](https://notmuchmail.org/manpages/notmuch-search-terms-7/) | Xapian search over a separately obtained mail store; shared field query language, threads, dates, attachments, tags | A small predictable field vocabulary and thread expansion; do not inherit English-only stemming as Turkish support |
| [mu](https://github.com/djcb/mu), [mu4e](https://www.djcbsoftware.nl/code/mu/mu4e/index.html) | Maildir indexing and efficient headers-to-message/thread navigation | Keep result summaries small and full reads explicit; avoid requiring Emacs/Xapian/Maildir for current CLI users |
| [himalaya](https://github.com/pimalaya/himalaya), [current backend sample](https://github.com/pimalaya/himalaya/blob/master/config.sample.toml) | Envelope/message distinction; pluggable IMAP, Maildir, notmuch and local pimdir backends | Local and remote reads share presentation concepts; missing local body remains explicit |
| [mbsync / isync](https://isync.sourceforge.io/mbsync.html) | UID state per mailbox pair; selectable direction; locking and durable sync state | Durable per-folder cursors, resume, bounded concurrency; avoid accidentally bringing two-way writes into a mirror |
| [OfflineIMAP](https://www.offlineimap.org/about/) | IMAP↔Maildir/IMAP synchronization, one-way or bidirectional | Interoperable mail export is useful later; another sync/auth/config process adds operations cost now |
| [lieer](https://github.com/gauteh/lieer) | Gmail API history with Maildir/notmuch integration and incremental label synchronization | A provider-native cursor and identity mapping; Gmail-specific history/labels do not transfer directly to Exchange |

Maildir preserves RFC822 interchange and makes bodies easy to inspect, but many
files, a second index engine, synchronization config and MIME payload copies are
unnecessary for a metadata/body/attachment-metadata mirror. Prefer one SQLite
database initially; offer an explicit Maildir/export path later if demanded.

## Storage and search

Use normalized metadata and address roles, stable conversation keys, extracted
plain text, attachment metadata and page/cursor state. Preserve separate message
instances: `InternetMessageId` can associate copies but is not a safe unique
deletion key. Deduplicate body blobs by hash rather than conflating sent/inbox
copies. Avoid indexing HTML markup, CSS, scripts, images or base64 attachment
content. Compress retained payloads and keep body output opt-in.

SQLite FTS5 supports relevance ranking and phrase/prefix queries. The delivered
store uses a **contentless word-token index**; compressed body blobs supply the
original text for application-generated snippets and Unicode-aware highlights.
There is no additional full plaintext column in the production database.
External-content FTS is an alternative when plaintext already exists in a table.
The English Porter tokenizer is unsuitable for Turkish. Trigram search supports infix
matching, with a minimum three-character full-text token; `detail=none` reduces
index size but longer MATCH tokens require a different approach such as indexed
LIKE. [FTS5](https://sqlite.org/fts5.html)

Normalize text and queries identically: Unicode case folding, `ı` to `i`, Unicode
decomposition, and combining-mark removal cover `İ/I/ı/i`, `ş/s`, `ğ/g`, `ü/u`,
`ö/o`, `ç/c`. Retain original text for display. Prefix `toplant*` finds suffix
forms but is not full linguistic stemming. Snowball's Turkish algorithm handles
noun/nominal-verb suffixes and consonant changes; even that is not a complete
morphological analyzer. [Turkish Snowball](https://snowballstem.org/algorithms/turkish/stemmer.html)

### Small reproducible lexical experiment

`python scripts/benchmark_search_engines.py` builds **10,000 synthetic rows**,
with repeated Turkish business phrases and unique identifiers, then executes
eight queries 30 times each. COUNT latency is warm in-process SQLite latency;
there is no CLI startup, HTTP, output serialization, or real-mail representativeness.
Run on the development machine on the review date:

| Index | Build seconds | Database bytes | COUNT p50 ms | COUNT p95 ms |
|---|---:|---:|---:|---:|
| Word tokens + prefix | 0.200 | 6,180,864 | 0.079 | 0.115 |
| Trigram, full detail | 0.276 | 12,529,664 | 0.620 | 0.772 |
| Trigram, no detail + LIKE | 0.232 | 6,119,424 | 1.324 | 2.030 |

All three matched folded queries such as `TOPLANTI`, `isbirligi`, `calisma`,
`poliçe`, and `güvenlik`. Only trigrams matched internal fragments `üşteri` and
`leşme`. The fixture therefore demonstrates a concrete recall benefit and a
size/latency tradeoff for substring search. It does not establish production
sizes, morphology precision, or mailbox p50/p95. Use the mailbox benchmark for
those claims.

Choose contentless word tokens for the initial store. The synthetic small-trigram
LIKE variant has an existing plaintext content table; that does not establish
the same storage footprint for a compressed-body store. Providing indexed LIKE
with accessible plaintext or SQL decompression adds another query/storage path.
Full-detail trigram inflated both this fixture and the independent implementation
experiment, while prefix search provides suffix-form recall with one small index.
Infix matching remains an explicit follow-up, not a delivered capability.

### Mailbox benchmark and fair scope

`scripts/benchmark_local.py` opens both databases read-only, emits only counts,
timings and byte totals, and never synchronizes. CLI timing uses fresh processes
against private temporary SQLite backups; backup creation is excluded. Saved
full/incremental sync reports are reduced to numeric aggregates, without folder
names, addresses, message IDs or provider cursor/error text.

```sh
umask 077
python scripts/benchmark_local.py \
  --legacy ~/.cache/outlook-cli/index.sqlite3 \
  --store ~/.cache/outlook-cli/mail.sqlite3 \
  --comparable-scope --repeats 20 --cli-repeats 3 \
  --first-full-report full-sync.json --incremental-report incremental-sync.json \
  > benchmark.json
```

The normal section measures the legacy scope and the new whole-mailbox prefix
search separately. Their counts and match modes differ; do not present that as a
comparison over identical corpora. `--comparable-scope` adds exact searches in
each legacy folder mapped by retained ID or unique Turkish-folded name, with 25
results per folder. It reports medians of folder/query p50/p95 measurements,
**not** the latency of a single search across their union. A missing or ambiguous
mapping makes scope comparison unknown and prevents partial-scope timing claims.
Matched folders still have different snapshot dates and indexed HTML/plain-text
representation; inspect counts and completeness before interpreting a speedup.

### Semantic search decision

Defer default embeddings. `sqlite-vec` can store vectors, partition them and
apply metadata filters, but it introduces an extension/model lifecycle and its
metadata filtering does not support every SQL operator.
[sqlite-vec](https://alexgarcia.xyz/sqlite-vec/features/vec0.html)

Calculated minimum storage for one float32 vector per message is
`messages × dimensions × 4`: 100,000 messages need **146.5 MiB at 384 dimensions**
or **293.0 MiB at 768 dimensions**, before table/model overhead. Four chunks per
message multiply that by four. This is a calculated floor, not an observed DB
size. Embedding throughput and Turkish relevance have not been measured with
this mailbox, so there is no evidence to justify making model downloads and
re-embedding mandatory. Lexical search provides exact people, domains, policy
numbers and reproducible counted filters. A later opt-in hybrid reranker should
first beat a labeled Turkish query set on recall@k and useful-find rate, while
reporting embedding build time, memory, model version, disk bytes and query p95.

## Operations and remaining ideas

Keep searches network-free. An explicit sync before a research session supplies
freshness without making every query depend on connectivity or authentication.
A timer helps frequently changing mail, but needs a user decision about background
access, token refresh and cadence. Do not install launchd automatically.

Store directories are owner-only and mail files must be 0600, including WAL/SHM
sidecars, lock/checkpoint files and temporary replacements. Permissions and
compression are not encryption. Retain the legacy index; build alongside it or
provide an explicit migration. A local purge must remove the selected store and
its sidecars without touching the remote mailbox, authentication or an unrelated
legacy index. Backups and SSD snapshots are outside logical file deletion.

The delivered implementation covers native REST full/delta/resume checkpoints,
folder discovery, compressed bodies, attachment metadata pagination, offline
search/read/thread/related-person/attachment navigation, and legacy preservation.
Mock tests cover pagination, cursor resets, sparse changes and moved/deleted
records. Real-account verification must use read-only enumeration and naturally
occurring changes; deliberate live mailbox moves or deletions are outside the
authorized validation scope.
Conservative optional suffix expansion is available, but complete Turkish
linguistic stemming and arbitrary infix matching are not delivered.

Remaining work, ordered by likely value:

1. Make supported Graph provisioning straightforward when the user chooses an
   Entra app registration. Continue evaluating long-running cursor expiry,
   tenant migrations and naturally occurring moves/deletions through read-only
   observations; use mocks for intentionally constructed mailbox changes.
2. Evaluate full Turkish linguistic stemming and infix recall on labeled queries;
   choose expansion/indexing only when recall gain outweighs false positives and
   measured disk/query cost. Keep current prefix and exact semantics predictable.
3. Add optional attachment-content extraction on requested downloads, with byte
   limits and provenance; never fetch all attachment binaries implicitly.
4. Offer explicit scheduled sync after cadence/consent choice; show freshness in
   every local result rather than automatically connecting during search.
5. Evaluate opt-in hybrid semantic reranking with Turkish relevance labels before
   adding it; model throughput, relevance and mailbox embedding size remain
   unmeasured, while the calculated minimum vector storage is recorded above.
6. Maildir/RFC822 export and application-level at-rest encryption if requested.
