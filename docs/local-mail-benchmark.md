# Local mailbox benchmark methodology

Real-account benchmark reports remain local and are not committed to this
repository. The reproducible synthetic lexical experiment is documented in
[design research](local-store-research.md). Synthetic results are not production
mailbox latency or storage claims.

`scripts/benchmark_local.py` measures the retained legacy query path and the new
local store without synchronization. It reports numeric counts, timings and byte
totals without mail content, addresses, message IDs, folder names or opaque
provider cursors. Account-level counts and measurements should remain private
unless the user explicitly chooses to publish them.

## Run privately

Use Python from an environment with this repository installed:

```sh
umask 077
python scripts/benchmark_local.py \
  --legacy ~/.cache/outlook-cli/index.sqlite3 \
  --store ~/.cache/outlook-cli/mail.sqlite3 \
  --comparable-scope --repeats 20 --cli-repeats 3 \
  > benchmark.json
```

The suite contains sixteen static Turkish business queries, including accented,
ASCII and uppercase variants, with 25 results per query. Warm timings include
scope/count/ranking/selected-row decoding/snippets and exclude JSON serialization.
Fresh CLI timings include process startup, imports, DB opening, querying and
serialization. JSON bytes per result are per-query averages; compare their
median separately from the distribution of individual message sizes.

Original databases open read-only. CLI runs use owner-only temporary committed
SQLite backups, removed automatically; copying is excluded from query timing.
Record Python/SQLite versions and scope before interpreting differences.

Legacy/new whole-store sections may have different scopes, snapshot dates,
HTML/plain-text representations, normalization and match modes. They must not be
presented as an identical-corpus speedup. `--comparable-scope` maps legacy folders
by retained ID or unique folded name and measures exact queries separately in
each matched folder. Reported aggregate p50/p95 values are medians across
folder/query measurements, not a combined union query. Missing or ambiguous
folder mapping fails the comparable-scope measurement.

## Full and incremental synchronization

Capture a full report during initial copying or an explicitly requested
`local sync --full` run. Normal sync on an already completed store is delta.
Raw sync JSON contains folder names/IDs and must remain private:

```sh
umask 077
outlook local sync --full --no-input --json > full-sync.json
outlook local sync --no-input --json > delta-sync.json
python scripts/benchmark_local.py \
  --legacy ~/.cache/outlook-cli/index.sqlite3 \
  --store ~/.cache/outlook-cli/mail.sqlite3 \
  --comparable-scope \
  --first-full-report full-sync.json --incremental-report delta-sync.json \
  > benchmark.json
```

Sync-report ingestion retains only numeric aggregates and enumerated modes.
For interrupted copies, sum active run durations and state whether manual pauses
are included. Successful GET counts exclude retries; response bytes are decoded
HTTP content rather than compressed wire bandwidth. A zero-change delta is not a
changed-mail workload benchmark. Validate deliberately constructed deletions,
moves, cursor expiry and throttling with mocks; live mailbox validation remains
read-only.

See [usage and operations](local-mail.md) and the separate synthetic
`scripts/benchmark_search_engines.py` experiment. Semantic model throughput and
Turkish relevance require an explicit separate evaluation; no embedding results
are implied by the lexical measurements.
