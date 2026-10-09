#!/usr/bin/env python3
"""Offline before/after mailbox benchmark with content-free JSON results.

Original databases open read-only. CLI runs use private SQLite snapshots, so the
legacy constructor cannot change the user's retained index. Snapshot creation is
excluded from query timing. This script never synchronizes or contacts Outlook.
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import sqlite3
import statistics
import subprocess
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

# Only static research terms are used. Even these are omitted from the report;
# query_id identifies the same ordinal in this public, reproducible suite.
QUERIES = (
    "toplantı", "toplanti", "TOPLANTI", "çalışma", "calisma", "ÇALIŞMA",
    "görüşme", "gorusme", "poliçe", "police", "sözleşme", "sozlesme",
    "güvenlik", "guvenlik", '"iş sürekliliği"', "yenileme",
)
ALLOWED_METRICS = (
    "elapsed_seconds", "requests", "request_count", "retry_count", "response_bytes",
    "selected_folders", "completed_folders", "remaining_folders", "partial",
)


def percentile(values, fraction):
    ordered = sorted(values)
    return round(ordered[max(0, math.ceil(fraction * len(ordered)) - 1)], 3)


def latency(values):
    return {"samples": len(values), "p50_ms": round(statistics.median(values), 3),
            "p95_ms": percentile(values, .95), "min_ms": round(min(values), 3),
            "max_ms": round(max(values), 3)} if values else None


def counts(values):
    return {"min": min(values), "max": max(values), "stable": len(set(values)) == 1} if values else None


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def safe_error(exc):
    # str(exc) can contain queries, file paths, SQL values or opaque provider IDs.
    return {"code": "benchmark_failure", "exception_type": type(exc).__name__}


def database_bytes(path):
    files = {suffix or "database": Path(str(path) + suffix).stat().st_size
             for suffix in ("", "-wal", "-shm") if Path(str(path) + suffix).is_file()}
    return {"files": files, "total": sum(files.values())}


def readonly_connection(path):
    db = sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True, timeout=30)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA query_only=ON")
    return db


def open_legacy(path):
    from outlook_cli.index_store import MailIndex
    # Exercise the actual 0.2 query implementation without its writable constructor.
    store = MailIndex.__new__(MailIndex)
    store.path, store.db = path, readonly_connection(path)
    return store


def measure_store(store, *, backend, repeats, limit, match_mode=None):
    from outlook_cli.pagination import Page
    from outlook_cli.serialization import to_json_envelope
    results, all_times, identities = [], [], []
    for position, query in enumerate(QUERIES, 1):
        options = {"backend": backend, "limit": limit}
        if match_mode is not None:
            options["match_mode"] = match_mode
        timings, returned, total, output_bytes, data_bytes = [], [], [], [], []
        row_ids = set()
        try:
            for _ in range(repeats):
                started = time.perf_counter_ns()
                rows, meta = store.query(query, **options)
                timings.append((time.perf_counter_ns() - started) / 1e6)
                returned.append(len(rows))
                total.append(meta.get("total_matches", len(rows)))
                row_ids = {row.get("id") for row in rows}
                envelope = to_json_envelope(Page(rows, meta=meta))
                output_bytes.append(len(envelope.encode("utf-8")) + 1)
                # Count data alone separately: envelope scope metadata is not a result.
                data_bytes.append(len(to_json_envelope(rows).encode("utf-8")))
            item = {"query_id": f"q{position:02d}", "ok": True,
                    "latency": latency(timings), "returned_count": counts(returned),
                    "total_matches": counts(total), "json_bytes": counts(output_bytes),
                    "json_bytes_per_result": round(statistics.median(output_bytes) / returned[-1], 1) if returned[-1] else None,
                    "data_envelope_bytes_per_result": round(statistics.median(data_bytes) / returned[-1], 1) if returned[-1] else None,
                    "scope_complete": bool(meta.get("complete")),
                    "result_complete": bool(meta.get("result_complete"))}
            all_times.extend(timings)
        except Exception as exc:  # noqa: BLE001 - never print potentially sensitive exception values
            item = {"query_id": f"q{position:02d}", "ok": False, "error": safe_error(exc)}
        results.append(item)
        identities.append(row_ids)
    return {"ok": all(row["ok"] for row in results), "latency": latency(all_times),
            "queries": results}, identities


def private_snapshot(source, target):
    """SQLite backup includes committed WAL pages and never changes source state."""
    target.parent.mkdir(parents=True, mode=0o700)
    descriptor = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    os.close(descriptor)
    original = readonly_connection(source)
    copy = sqlite3.connect(target)
    try:
        original.backup(copy)
    finally:
        original.close()
        copy.close()
    os.chmod(target, 0o600)


def telemetry(stderr):
    for line in stderr.splitlines():
        try:
            value = json.loads(line)
        except (ValueError, TypeError):
            continue
        if isinstance(value, dict) and isinstance(value.get("profile"), dict):
            profile = value["profile"]
            return {key: profile[key] for key in ("request_count", "retry_count", "response_bytes")
                    if key in profile and numeric(profile[key])}
    return None


def measure_cli(executable, cache, config, *, group, backend, repeats, limit, match_mode=None, timeout=60):
    env = os.environ.copy()
    env.update(OUTLOOK_CLI_CACHE=str(cache), OUTLOOK_CLI_CONFIG=str(config),
               OUTLOOK_ENABLE_COMMANDS=group, OUTLOOK_ACCOUNT="default")
    # No tokens or auth metadata are copied into these isolated directories.
    results, all_times = [], []
    for position, query in enumerate(QUERIES, 1):
        command = [executable, "--no-input", "--profile", group, "search", query,
                   "--account", "default", "--backend", backend, "--limit", str(limit), "--json"]
        if match_mode is not None:
            command.extend(("--match", match_mode))
        timings, returned, total, output_bytes, request_counts = [], [], [], [], []
        try:
            for _ in range(repeats):
                started = time.perf_counter_ns()
                process = subprocess.run(command, capture_output=True, env=env, timeout=timeout, check=False)
                elapsed = (time.perf_counter_ns() - started) / 1e6
                envelope = json.loads(process.stdout)
                if process.returncode or envelope.get("ok") is not True:
                    error_code = str((envelope.get("error") or {}).get("code", "cli_failure"))
                    if not re.fullmatch(r"[a-z_]{1,60}", error_code):
                        error_code = "invalid_error_code"
                    results.append({"query_id": f"q{position:02d}", "ok": False,
                                    "exit_code": process.returncode, "error": {"code": error_code}})
                    break
                rows = envelope.get("data")
                if not isinstance(rows, list):
                    raise TypeError("Unexpected search result shape")
                profile = telemetry(process.stderr)
                if profile and profile.get("request_count", 0):
                    raise RuntimeError("Unexpected network request in offline benchmark")
                timings.append(elapsed)
                returned.append(len(rows))
                total.append((envelope.get("meta") or {}).get("total_matches", len(rows)))
                output_bytes.append(len(process.stdout))
                request_counts.append(profile.get("request_count") if profile else None)
            else:
                results.append({"query_id": f"q{position:02d}", "ok": True,
                                "latency": latency(timings), "returned_count": counts(returned),
                                "total_matches": counts(total), "json_bytes": counts(output_bytes),
                                "json_bytes_per_result": round(statistics.median(output_bytes) / returned[-1], 1) if returned[-1] else None,
                                "recorded_network_requests": sum(request_counts) if all(x is not None for x in request_counts) else None})
                all_times.extend(timings)
        except Exception as exc:  # noqa: BLE001 - content-free failure reports
            results.append({"query_id": f"q{position:02d}", "ok": False, "error": safe_error(exc)})
    return {"ok": all(row["ok"] for row in results), "latency": latency(all_times), "queries": results}


def ingest_sync(path, label):
    try:
        report = json.loads(path.read_text())
        meta = report.get("meta") or {}
        outcomes = report.get("data") or []
        if not isinstance(outcomes, list):
            raise TypeError("Expected sync result list")
        result = {"label": label, "ok": report.get("ok") is True,
                  **{key: meta[key] for key in ALLOWED_METRICS
                     if key in meta and (numeric(meta[key]) or isinstance(meta[key], bool))},
                  "folder_outcomes": len(outcomes),
                  "successful_folder_outcomes": sum(item.get("ok") is True for item in outcomes),
                  "failed_folder_outcomes": sum(item.get("ok") is not True for item in outcomes),
                  "modes": dict(Counter(item.get("mode") if item.get("mode") in
                                        {"full", "delta", "snapshot", "full_metadata"} else "unknown"
                                        for item in outcomes)),
                  **{key: sum(item.get(key, 0) for item in outcomes if numeric(item.get(key, 0)))
                     for key in ("records", "removed", "pages")}}
        # Never echo error.message: a provider URL can contain an opaque cursor.
        if not result["ok"]:
            result["error"] = {"code": "sync_report_incomplete"}
        return result
    except Exception as exc:  # noqa: BLE001 - content-free failure reports
        return {"label": label, "ok": False, "error": safe_error(exc)}


def arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--legacy", type=Path, required=True, help="Existing index.sqlite3; opened read-only")
    parser.add_argument("--store", type=Path, required=True, help="New mail.sqlite3; opened read-only")
    parser.add_argument("--repeats", type=int, default=20)
    parser.add_argument("--cli-repeats", type=int, default=3, help="Set 0 to skip subprocess and private DB snapshots")
    parser.add_argument("--outlook", default=shutil.which("outlook") or "outlook", help="Installed CLI executable")
    parser.add_argument("--backend", choices=("rest", "graph"), default="rest")
    parser.add_argument("--match", choices=("exact", "prefix", "stem"), default="prefix")
    parser.add_argument("--limit", type=int, default=25)
    parser.add_argument("--timeout", type=int, default=60, help="Seconds per offline CLI invocation")
    parser.add_argument("--first-full-report", type=Path, help="Read saved full-sync JSON; do not run sync")
    parser.add_argument("--incremental-report", type=Path, help="Read saved incremental-sync JSON; do not run sync")
    parser.add_argument("--sync-report", type=Path, action="append", default=[], help="Additional saved sync JSON")
    args = parser.parse_args()
    if args.repeats < 1 or args.cli_repeats < 0 or not 1 <= args.limit <= 100000 or args.timeout < 1:
        parser.error("repeats/timeout must be positive, cli-repeats nonnegative, limit 1..100000")
    return args


def main():
    args = arguments()
    result = {"offline": True, "query_suite": "turkish_business_v1", "query_count": len(QUERIES),
              "limit": args.limit, "backend": args.backend, "new_match_mode": args.match,
              "warm_query_includes": "scope/count/ranking/selected-row decoding/snippets; excludes JSON serialization",
              "cli_includes": "fresh Python process/imports/DB open/query/JSON serialization",
              "cli_snapshots": "private committed SQLite backups; copy cost excluded; original databases untouched",
              "database_bytes": {"legacy": database_bytes(args.legacy), "new": database_bytes(args.store)}}
    legacy = new = None
    failures = []
    try:
        from outlook_cli.mail_store import MailStore
        legacy, new = open_legacy(args.legacy), MailStore(args.store, readonly=True)
        result["legacy_in_process"], old_ids = measure_store(legacy, backend=args.backend, repeats=args.repeats, limit=args.limit)
        result["new_in_process"], new_ids = measure_store(new, backend=args.backend, repeats=args.repeats, limit=args.limit, match_mode=args.match)
        old_scope = legacy.db.execute("SELECT count(*),count(DISTINCT folder_id) FROM messages WHERE backend=?", (args.backend,)).fetchone()
        status = new.status(args.backend)
        result["scope"] = {"legacy_message_count": old_scope[0], "legacy_folder_count": old_scope[1],
                           "new_message_count": status.get("message_count"), "new_folder_count": status.get("folder_count"),
                           "new_whole_mailbox_complete": status.get("whole_mailbox_complete"),
                           "counts_directly_comparable": old_scope[0] == status.get("message_count") and old_scope[1] == status.get("folder_count")}
        result["correctness"] = [{"query_id": old["query_id"],
                                  "legacy_total_matches": old.get("total_matches"),
                                  "new_total_matches": current.get("total_matches"),
                                  "shared_returned_id_count": len(a.intersection(b)),
                                  "returned_id_set_equal": a == b,
                                  "interpretation": "scope/text normalization/match mode/ranking can differ"}
                                 for old, current, a, b in zip(result["legacy_in_process"]["queries"], result["new_in_process"]["queries"], old_ids, new_ids)]
        if args.cli_repeats:
            # Snapshots contain sensitive mail: private directories/files and automatic cleanup.
            with tempfile.TemporaryDirectory(prefix="outlook-offline-benchmark-") as scratch:
                base = Path(scratch)
                old_cache, new_cache, config = base / "legacy", base / "new", base / "config"
                config.mkdir(mode=0o700)
                private_snapshot(args.legacy, old_cache / "index.sqlite3")
                private_snapshot(args.store, new_cache / "mail.sqlite3")
                result["legacy_cli"] = measure_cli(args.outlook, old_cache, config, group="index", backend=args.backend,
                                                   repeats=args.cli_repeats, limit=args.limit, timeout=args.timeout)
                result["new_cli"] = measure_cli(args.outlook, new_cache, config, group="local", backend=args.backend,
                                                repeats=args.cli_repeats, limit=args.limit, match_mode=args.match, timeout=args.timeout)
    except Exception as exc:  # noqa: BLE001 - close both stores and emit content-free failure
        failures.append(safe_error(exc))
    finally:
        if legacy:
            legacy.close()
        if new:
            new.close()
    reports = []
    for label, path in (("first_full", args.first_full_report), ("incremental", args.incremental_report)):
        if path:
            reports.append(ingest_sync(path, label))
    reports.extend(ingest_sync(path, f"additional_{number}") for number, path in enumerate(args.sync_report, 1))
    result["sync_reports"] = reports
    okay = not failures and all(value.get("ok", True) for value in result.values() if isinstance(value, dict)) and all(report["ok"] for report in reports)
    if failures:
        result["errors"] = failures
    print(json.dumps({"ok": okay, "schema_version": "1", "data": result,
                      "meta": {"contains_mail_content": False, "contains_addresses": False,
                               "contains_message_ids": False, "no_sync_performed": True}}, ensure_ascii=False, indent=2))
    return 0 if okay else 1


if __name__ == "__main__":
    raise SystemExit(main())
