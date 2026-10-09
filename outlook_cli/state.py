"""Transactional profile state; display numbers are permanent references."""
from __future__ import annotations

from contextlib import contextmanager
import json
import os
from pathlib import Path
import sqlite3

from .exceptions import AccountError
from .locking import atomic_write_json as atomic_json_write, file_lock


@contextmanager
def locked_file(path: Path):
    """Shared bounded lock for a profile file's read/modify/write transaction."""
    with file_lock(path.with_suffix(path.suffix + ".lock")):
        yield


class IDStore:
    """SQLite allocation with no eviction or reuse, including legacy numbers."""
    def __init__(self, legacy_path: Path):
        self.legacy_path = legacy_path
        self.path = legacy_path.with_suffix(".sqlite3")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        # Initialization is also serialized: concurrent first-use must migrate once.
        with locked_file(self.path):
            with self._connect() as db:
                db.execute("CREATE TABLE IF NOT EXISTS identities (num INTEGER PRIMARY KEY AUTOINCREMENT, real_id TEXT NOT NULL UNIQUE)")
                db.execute("CREATE TABLE IF NOT EXISTS aliases (num INTEGER PRIMARY KEY, target TEXT NOT NULL)")
                db.execute("CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
                if db.execute("SELECT value FROM metadata WHERE key='legacy_migrated'").fetchone() is None:
                    legacy = self._read_legacy()
                    try:
                        db.executemany("INSERT INTO identities(num,real_id) VALUES (?,?)", legacy)
                    except sqlite3.IntegrityError as exc:
                        raise AccountError("Conflicting legacy display numbers; ID map was preserved for repair.") from exc
                    db.execute("INSERT INTO metadata VALUES ('legacy_migrated','1')")
            os.chmod(self.path, 0o600)

    @contextmanager
    def _connect(self):
        db = sqlite3.connect(self.path, timeout=30)
        db.execute("PRAGMA busy_timeout=30000")
        try:
            with db:
                yield db
        finally:
            db.close()

    def _read_legacy(self):
        if not self.legacy_path.exists():
            return []
        try:
            value = json.loads(self.legacy_path.read_text())
            if not isinstance(value, dict):
                raise ValueError("expected object")
            pairs = []
            for key, real_id in value.items():
                if not str(key).isdigit() or int(key) <= 0 or not isinstance(real_id, str) or not real_id:
                    raise ValueError("invalid display number or ID")
                pairs.append((int(key), real_id))
            return pairs
        except (OSError, ValueError) as exc:
            raise AccountError(f"Cannot migrate corrupt ID map {self.legacy_path}; original file preserved, no numbers reassigned.") from exc

    def allocate(self, real_ids: list[str]) -> list[int]:
        if not real_ids:
            return []
        with self._connect() as db:
            existing = {row[0]: row[1] for real_id in dict.fromkeys(real_ids)
                        if (row := db.execute("SELECT real_id,num FROM identities WHERE real_id=?", (real_id,)).fetchone())}
        if len(existing) == len(set(real_ids)):
            return [existing[real_id] for real_id in real_ids]
        result = []
        with self._connect() as db:
            db.execute("BEGIN IMMEDIATE")
            for real_id in real_ids:
                row = db.execute("SELECT num FROM identities WHERE real_id=?", (real_id,)).fetchone()
                if row is None:
                    row = (db.execute("INSERT INTO identities(real_id) VALUES (?)", (real_id,)).lastrowid,)
                result.append(row[0])
        return result

    def resolve(self, number: str) -> str | None:
        if not number.isdigit():
            return None
        with self._connect() as db:
            row = db.execute("SELECT COALESCE(a.target,i.real_id) FROM identities i LEFT JOIN aliases a ON a.num=i.num WHERE i.num=?", (int(number),)).fetchone()
        return row[0] if row else None

    def snapshot(self) -> dict[str, str]:
        with self._connect() as db:
            return {str(num): real_id for num, real_id in db.execute("SELECT i.num,COALESCE(a.target,i.real_id) FROM identities i LEFT JOIN aliases a ON a.num=i.num")}

    def replace(self, old_id: str, new_id: str) -> int:
        """Preserve the original display reference across provider ID changes."""
        with self._connect() as db:
            db.execute("BEGIN IMMEDIATE")
            old = db.execute("SELECT num FROM identities WHERE real_id=?", (old_id,)).fetchone()
            existing = db.execute("SELECT num FROM identities WHERE real_id=?", (new_id,)).fetchone()
            if old_id == new_id:
                return old[0] if old else db.execute("INSERT INTO identities(real_id) VALUES (?)", (new_id,)).lastrowid
            if old and existing:
                # Both references already exist. Keep the old one as an alias,
                # represented in a separate table, without recycling either num.
                db.execute("UPDATE aliases SET target=? WHERE target=?", (new_id, old_id))
                db.execute("INSERT OR REPLACE INTO aliases VALUES (?,?)", (old[0], new_id))
                return old[0]
            if old:
                db.execute("UPDATE identities SET real_id=? WHERE num=?", (new_id, old[0]))
                db.execute("UPDATE aliases SET target=? WHERE target=?", (new_id, old_id))
                return old[0]
            return existing[0] if existing else db.execute("INSERT INTO identities(real_id) VALUES (?)", (new_id,)).lastrowid
