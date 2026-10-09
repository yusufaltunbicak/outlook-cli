#!/usr/bin/env python3
"""Synthetic FTS alternatives benchmark; never opens tokens or a mailbox."""
from __future__ import annotations

import json
import random
import sqlite3
import statistics
import tempfile
import time
import unicodedata
from pathlib import Path


def fold(value: str) -> str:
    return "".join(
        char for char in unicodedata.normalize("NFKD", value.casefold().replace("ı", "i"))
        if not unicodedata.combining(char)
    )


def main() -> None:
    random.seed(11)
    sentences = [
        "Toplantılarımız ve işbirliği için çalışma planı güncellendi.",
        "Müşterinin poliçesi yenilendi, görüşmeler sürüyor.",
        "Ödeme ve müşteri iletişimi hakkında rapor.",
        "Sözleşmelerin değerlendirmesi ve güvenlik denetimleri.",
    ]
    queries = ["toplantı", "TOPLANTI", "isbirligi", "calisma", "üşteri", "poliçe", "güvenlik", "leşme"]
    measurements = {}
    for kind, tokenizer, detail in [
        ("word", "unicode61", "full"),
        ("trigram", "trigram", "full"),
        ("trigram_small", "trigram", "none"),
    ]:
        with tempfile.TemporaryDirectory(prefix="outlook-fts-synthetic-") as scratch:
            path = Path(scratch) / "benchmark.sqlite3"
            db = sqlite3.connect(path)
            db.execute("CREATE TABLE texts(text TEXT NOT NULL)")
            db.execute(
                "CREATE VIRTUAL TABLE ft USING fts5(text,content='texts',"
                f"content_rowid='rowid',tokenize='{tokenizer}',detail={detail})"
            )
            started = time.perf_counter()
            for row in range(10000):
                text = fold(" ".join([sentences[row % 4]] * 8) + f" dosya_{row} " + str(random.getrandbits(64)))
                db.execute("INSERT INTO texts VALUES(?)", (text,))
            db.execute("INSERT INTO ft(ft) VALUES('rebuild')")
            db.commit()
            build_s = time.perf_counter() - started
            latencies, counts = [], {}
            for query in queries:
                normalized = fold(query)
                if kind == "trigram_small":
                    sql, argument = "SELECT COUNT(*) FROM ft WHERE text LIKE ?", "%" + normalized + "%"
                else:
                    sql = "SELECT COUNT(*) FROM ft WHERE ft MATCH ?"
                    argument = '"' + normalized + '"' + ("*" if kind == "word" else "")
                counts[query] = db.execute(sql, (argument,)).fetchone()[0]
                for _ in range(30):
                    started_ns = time.perf_counter_ns()
                    db.execute(sql, (argument,)).fetchone()
                    latencies.append((time.perf_counter_ns() - started_ns) / 1e6)
            db.close()
            latencies.sort()
            measurements[kind] = {
                "rows": 10000,
                "build_s": round(build_s, 3),
                "db_bytes": path.stat().st_size,
                "search_count_p50_ms": round(statistics.median(latencies), 3),
                "search_count_p95_ms": round(latencies[int(.95 * (len(latencies) - 1))], 3),
                "counts": counts,
            }
    print(json.dumps({"synthetic": True, "sqlite_version": sqlite3.sqlite_version,
                      "measurements": measurements}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
