"""Private, offline mail storage with compressed bodies and resumable folder generations.

The legacy index is deliberately independent. FTS stores tokens, not another copy
of the HTML/text; only selected search results decompress bodies for snippets.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shlex
import sqlite3
import stat
import time
import unicodedata
import uuid
import zlib
from datetime import datetime, timezone
from pathlib import Path

from bs4 import BeautifulSoup

from .exceptions import OutlookCliError, ResourceNotFoundError

SCHEMA_VERSION = 1
_QUERY_FIELDS = {"from", "to", "person", "domain", "folder", "after", "before", "has", "thread", "conversation"}
# Conservative query expansion, not a full Turkish morphological analyser.
_SUFFIXES = tuple(sorted(("larindan", "lerinden", "larinin", "lerinin", "lardan", "lerden", "larina", "lerine", "larin", "lerin", "lari", "leri", "nin", "nun", "dan", "den", "lar", "ler", "nda", "nde", "yla", "yle", "da", "de"), key=len, reverse=True))


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def fold(text: str) -> str:
    """Turkish-safe accent/case folding; stored display text is untouched."""
    decomposed = unicodedata.normalize("NFKD", str(text).casefold().replace("ı", "i"))
    return "".join(char for char in decomposed if not unicodedata.combining(char))


def _stem(word: str) -> str:
    for suffix in _SUFFIXES:
        if word.endswith(suffix) and len(word) - len(suffix) >= 4:
            return word[:-len(suffix)]
    return word


def body_text(body: str, body_type: str) -> str:
    if body_type.casefold() != "html":
        return body
    soup = BeautifulSoup(body, "html.parser")
    for node in soup(["script", "style", "head", "noscript"]):
        node.decompose()
    return soup.get_text("\n", strip=True)


def _dump(value) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=lambda x: x.isoformat())


def _pack(value) -> bytes:
    return zlib.compress(_dump(value).encode("utf-8"), 6)


def _unpack(value: bytes) -> dict:
    return json.loads(zlib.decompress(value))


def _received(value) -> str:
    if not value:
        return ""
    parsed = value if isinstance(value, datetime) else datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    return parsed.replace(tzinfo=parsed.tzinfo or timezone.utc).astimezone(timezone.utc).isoformat()


def _date(value) -> str:
    try:
        return _received(value)
    except (ValueError, TypeError) as exc:
        raise ValueError("Expected an ISO date or datetime") from exc


def _terms(query: str) -> tuple[list[tuple[str, bool]], dict]:
    """Parse literal words/phrases and a small, predictable field vocabulary."""
    pattern = re.compile(r'''(?:(?P<field>[A-Za-z]+):)?(?P<value>"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|[^\s"']+)''')
    terms, fields, end = [], {}, 0
    for match in pattern.finditer(query):
        if query[end:match.start()].strip():
            raise ValueError("Unclosed quotation in local search query")
        end = match.end()
        raw = match.group("value")
        quoted = raw.startswith(('"', "'"))
        value = shlex.split(raw)[0] if quoted else raw
        field = (match.group("field") or "").lower()
        if field and field in _QUERY_FIELDS:
            key = {"from": "sender", "to": "recipient", "thread": "conversation"}.get(field, field)
            if field == "has":
                if value.casefold() not in ("attachment", "attachments"):
                    raise ValueError("has: supports attachment or attachments")
                fields["has_attachments"] = True
            else:
                fields[key] = value
        elif field:
            raise ValueError(f"Unknown local query field: {field}")
        elif value:
            terms.append((value, quoted or " " in value))
    if query[end:].strip():
        raise ValueError("Unclosed quotation in local search query")
    return terms, fields


def _fts_query(terms: list[tuple[str, bool]], mode: str) -> tuple[str, list[str]]:
    if mode not in ("exact", "prefix", "stem"):
        raise ValueError("match_mode must be exact, prefix or stem")
    clauses, highlights = [], []
    for raw, phrase in terms:
        term = fold(raw)
        # FTS treats punctuation as token separators, and literal quotes prevent
        # user input from becoming an FTS operator or column expression.
        words = re.findall(r"\w+", term, re.UNICODE)
        if not words:
            continue
        if phrase:
            clauses.append('"' + " ".join(words) + '"')
            highlights.append(term)
        else:
            for word in words:
                expanded = _stem(word) if mode == "stem" else word
                clauses.append('"' + expanded + '"' + ("*" if mode != "exact" else ""))
                highlights.append(expanded)
    return " AND ".join(clauses), highlights


def _snippet(text: str, terms: list[str], max_chars: int = 240) -> tuple[str, str]:
    """Bounded snippets highlight original spelling with plain [[...]] markers."""
    text = re.sub(r"\s+", " ", text).strip()
    folded_text = fold(text)
    # Turkish case/diacritic folding preserves codepoint positions. Only build
    # the slower mapping for decomposed accents or expanding Unicode letters.
    offsets = None
    if len(folded_text) != len(text):
        offsets = [index for index, char in enumerate(text) for _ in fold(char)]
    matches = []
    for term in terms:
        if not term:
            continue
        start = folded_text.find(term)
        if start >= 0:
            stop = start + len(term)
            # Prefix/stem matching highlights the complete word.
            while stop < len(folded_text) and folded_text[stop].isalnum():
                stop += 1
            matches.append((offsets[start], offsets[stop - 1] + 1) if offsets is not None else (start, stop))
    focus = min((start for start, _ in matches), default=0)
    start = max(0, focus - max_chars // 3)
    stop = min(len(text), start + max_chars)
    if start:
        boundary = text.find(" ", start, min(start + 24, stop))
        if boundary >= 0 and boundary < focus:
            start = boundary + 1
    leading, trailing = ("…" if start else ""), ("…" if stop < len(text) else "")
    snippet = leading + text[start:stop] + trailing
    spans = sorted({(max(left, start), min(right, stop)) for left, right in matches if right > start and left < stop})
    merged = []
    for left, right in spans:
        if merged and left <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(right, merged[-1][1]))
        else:
            merged.append((left, right))
    pieces, position = [leading], start
    for left, right in merged:
        pieces.extend((text[position:left], "[[", text[left:right], "]]"))
        position = right
    pieces.extend((text[position:stop], trailing))
    return snippet, "".join(pieces)


class MailStore:
    def __init__(self, path: Path, *, readonly: bool = False):
        self.path = Path(path)
        self.readonly = readonly
        if readonly:
            self.db = sqlite3.connect(self.path.resolve().as_uri() + "?mode=ro", uri=True, timeout=30)
            self.db.row_factory = sqlite3.Row
            self.db.execute("PRAGMA query_only=ON")
            if self.db.execute("PRAGMA user_version").fetchone()[0] != SCHEMA_VERSION:
                self.db.close()
                raise OutlookCliError("Unsupported mail store schema")
            return
        self.path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        os.chmod(self.path.parent, 0o700)
        descriptor = os.open(self.path, os.O_RDWR | os.O_CREAT | getattr(os, "O_NOFOLLOW", 0), 0o600)
        try:
            if not stat.S_ISREG(os.fstat(descriptor).st_mode):
                raise OutlookCliError("Mail storage must be a regular private file")
            os.fchmod(descriptor, 0o600)
        finally:
            os.close(descriptor)
        self.db = sqlite3.connect(self.path, timeout=30)
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.execute("PRAGMA secure_delete=ON")
        version = self.db.execute("PRAGMA user_version").fetchone()[0]
        if version not in (0, SCHEMA_VERSION):
            self.db.close()
            raise OutlookCliError("Unsupported mail store schema; keep this file and upgrade the CLI")
        self.db.executescript("""
        CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS folders (
            backend TEXT NOT NULL, id TEXT NOT NULL, name TEXT NOT NULL,
            active_generation TEXT, pending_generation TEXT, pending_full INTEGER NOT NULL DEFAULT 0,
            pending_url TEXT, cursor TEXT, seed_done INTEGER NOT NULL DEFAULT 0,
            snapshot_mode INTEGER NOT NULL DEFAULT 0,
            options_signature TEXT NOT NULL DEFAULT '', include_body INTEGER NOT NULL DEFAULT 1,
            synced_at TEXT, complete INTEGER NOT NULL DEFAULT 0, error TEXT,
            PRIMARY KEY(backend,id));
        CREATE TABLE IF NOT EXISTS bodies (
            hash TEXT PRIMARY KEY, original BLOB NOT NULL, plain BLOB, body_type TEXT NOT NULL,
            original_bytes INTEGER NOT NULL, text_bytes INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS messages (
            record_id INTEGER PRIMARY KEY, backend TEXT NOT NULL, id TEXT NOT NULL, folder_id TEXT NOT NULL,
            generation TEXT NOT NULL, received TEXT NOT NULL, conversation_id TEXT NOT NULL,
            has_attachments INTEGER NOT NULL, body_hash TEXT REFERENCES bodies(hash), metadata BLOB NOT NULL,
            UNIQUE(backend,folder_id,generation,id));
        CREATE INDEX IF NOT EXISTS mail_identity ON messages(backend,id);
        CREATE INDEX IF NOT EXISTS mail_scope ON messages(backend,folder_id,generation,received);
        CREATE INDEX IF NOT EXISTS mail_conversation ON messages(backend,conversation_id,received);
        CREATE TABLE IF NOT EXISTS addresses (
            message_rowid INTEGER NOT NULL REFERENCES messages(record_id) ON DELETE CASCADE,
            address TEXT NOT NULL, name TEXT NOT NULL, domain TEXT NOT NULL, role TEXT NOT NULL,
            PRIMARY KEY(message_rowid,address,role));
        CREATE INDEX IF NOT EXISTS mail_address ON addresses(address,role,message_rowid);
        CREATE INDEX IF NOT EXISTS mail_domain ON addresses(domain,message_rowid);
        CREATE TABLE IF NOT EXISTS attachments (
            message_rowid INTEGER NOT NULL REFERENCES messages(record_id) ON DELETE CASCADE,
            id TEXT NOT NULL, name TEXT NOT NULL, content_type TEXT NOT NULL, size INTEGER NOT NULL,
            is_inline INTEGER NOT NULL, PRIMARY KEY(message_rowid,id));
        CREATE VIRTUAL TABLE IF NOT EXISTS mail_fts USING fts5(
            subject,people,body,attachments,content='',tokenize='unicode61 remove_diacritics 2',prefix='2 3');
        """)
        self.db.execute(f"PRAGMA user_version={SCHEMA_VERSION}")
        self.db.commit()
        self._private_files()

    def _private_files(self):
        if self.readonly:
            return
        for path in (self.path, Path(str(self.path) + "-wal"), Path(str(self.path) + "-shm")):
            if path.exists() and not path.is_symlink():
                os.chmod(path, 0o600)

    def close(self):
        self.db.close()
        self._private_files()

    def folder_state(self, backend: str, folder_id: str) -> dict:
        row = self.db.execute("SELECT * FROM folders WHERE backend=? AND id=?", (backend, folder_id)).fetchone()
        return dict(row) if row else {}

    def begin_sync(self, backend: str, folder_id: str, name: str, *, full: bool,
                   include_body: bool = True, options_signature: str = "", restart: bool = False) -> dict:
        state = self.folder_state(backend, folder_id)
        if not restart and state.get("pending_generation") and state.get("options_signature") == options_signature and bool(state["include_body"]) == include_body:
            return state
        with self.db:
            if state.get("pending_full") and state.get("pending_generation"):
                self._delete_where("backend=? AND folder_id=? AND generation=?", (backend, folder_id, state["pending_generation"]))
            generation = uuid.uuid4().hex if full or not state.get("active_generation") else state["active_generation"]
            self.db.execute("""INSERT INTO folders(backend,id,name,pending_generation,pending_full,include_body,options_signature)
                VALUES(?,?,?,?,?,?,?) ON CONFLICT(backend,id) DO UPDATE SET name=excluded.name,
                pending_generation=excluded.pending_generation,pending_full=excluded.pending_full,pending_url=NULL,
                seed_done=0,snapshot_mode=0,include_body=excluded.include_body,options_signature=excluded.options_signature,error=NULL""",
                (backend, folder_id, name, generation, int(full or not state.get("active_generation")), int(include_body), options_signature))
        self._private_files()
        return self.folder_state(backend, folder_id)

    def _body(self, body_hash: str | None, *, original=False) -> str:
        if not body_hash:
            return ""
        row = self.db.execute("SELECT original,plain FROM bodies WHERE hash=?", (body_hash,)).fetchone()
        compressed = row["original"] if original or row["plain"] is None else row["plain"]
        return zlib.decompress(compressed).decode("utf-8")

    @staticmethod
    def _fts_values(record: dict, plain: str) -> tuple[str, str, str, str]:
        people = [record.get("sender") or {}] + record.get("to", []) + record.get("cc", []) + record.get("bcc", [])
        names = " ".join(f"{person.get('name', '')} {person.get('address', '')}" for person in people)
        attachment_names = " ".join(item.get("name", "") for item in record.get("attachments", []))
        return tuple(fold(value) for value in (record.get("subject", ""), names, plain or record.get("preview", ""), attachment_names))

    def _delete_where(self, where: str, args):
        rows = self.db.execute(f"SELECT record_id,body_hash,metadata FROM messages WHERE {where}", args).fetchall()
        for row in rows:
            values = self._fts_values(_unpack(row["metadata"]), self._body(row["body_hash"]))
            self.db.execute("INSERT INTO mail_fts(mail_fts,rowid,subject,people,body,attachments) VALUES('delete',?,?,?,?,?)", (row["record_id"], *values))
            self.db.execute("DELETE FROM messages WHERE record_id=?", (row["record_id"],))

    def _put(self, backend: str, folder_id: str, generation: str, message: dict):
        record = dict(message)
        record.pop("display_num", None)
        body = str(record.pop("body", "") or "")
        body_type = str(record.get("body_type", "Text"))
        digest, plain = None, ""
        if body:
            digest = hashlib.sha256((body_type.casefold() + "\0" + body).encode("utf-8")).hexdigest()
            known = self.db.execute("SELECT hash FROM bodies WHERE hash=?", (digest,)).fetchone()
            if known:
                plain = self._body(digest)
            else:
                plain = body_text(body, body_type)
                raw, text = body.encode("utf-8"), plain.encode("utf-8")
                self.db.execute("INSERT INTO bodies VALUES(?,?,?,?,?,?)", (digest, zlib.compress(raw, 6), zlib.compress(text, 6) if body_type.casefold() == "html" else None, body_type, len(raw), len(text)))
        # Store metadata only, never base64 payloads or credential/debug fields.
        record["attachments"] = [{key: item[key] for key in ("id", "name", "content_type", "size", "is_inline") if key in item} for item in record.get("attachments", [])]
        for key in ("token", "access_token", "cookie", "content_bytes", "ContentBytes"):
            record.pop(key, None)
        record["received"] = _received(record.get("received"))
        self._delete_where("backend=? AND folder_id=? AND generation=? AND id=?", (backend, folder_id, generation, record["id"]))
        result = self.db.execute("""INSERT INTO messages(backend,id,folder_id,generation,received,conversation_id,has_attachments,body_hash,metadata)
            VALUES(?,?,?,?,?,?,?,?,?)""", (backend, record["id"], folder_id, generation, record["received"], record.get("conversation_id") or "", int(bool(record.get("has_attachments") or record["attachments"])), digest, _pack(record)))
        rowid = result.lastrowid
        self.db.execute("INSERT INTO mail_fts(rowid,subject,people,body,attachments) VALUES(?,?,?,?,?)", (rowid, *self._fts_values(record, plain)))
        people = [(record.get("sender") or {}, "from")]
        for role in ("to", "cc", "bcc"):
            people += [(person, role) for person in record.get(role, [])]
        self.db.executemany("INSERT OR IGNORE INTO addresses VALUES(?,?,?,?,?)", [(rowid, fold(person.get("address", "")), fold(person.get("name", "")), fold(person.get("address", "")).partition("@")[2], role) for person, role in people if person.get("address")])
        self.db.executemany("INSERT OR IGNORE INTO attachments VALUES(?,?,?,?,?,?)", [(rowid, str(item.get("id") or f"metadata-{index}"), str(item.get("name", "")), str(item.get("content_type", "")), int(item.get("size") or 0), int(bool(item.get("is_inline")))) for index, item in enumerate(record["attachments"])])

    def _remove_moved(self, backend: str, folder_id: str, generation: str):
        self._delete_where("backend=? AND folder_id<>? AND id IN (SELECT id FROM messages WHERE backend=? AND folder_id=? AND generation=?) AND generation IN (SELECT active_generation FROM folders WHERE backend=?)", (backend, folder_id, backend, folder_id, generation, backend))

    def apply_page(self, backend: str, folder_id: str, name: str, messages: list[dict], *,
                   removed=(), next_url: str | None = None, cursor: str | None = None,
                   complete: bool = False, seed_done: bool | None = None,
                   snapshot_mode: bool | None = None):
        state = self.folder_state(backend, folder_id)
        if not state.get("pending_generation"):
            raise OutlookCliError("begin_sync is required before applying mail pages")
        generation = state["pending_generation"]
        with self.db:
            for identity in removed:
                self._delete_where("backend=? AND folder_id=? AND generation=? AND id=?", (backend, folder_id, generation, identity))
            for message in messages:
                self._put(backend, folder_id, generation, message)
            if complete:
                self._remove_moved(backend, folder_id, generation)
                old = state.get("active_generation")
                if state["pending_full"] and old and old != generation:
                    self._delete_where("backend=? AND folder_id=? AND generation=?", (backend, folder_id, old))
                self.db.execute("""UPDATE folders SET name=?,active_generation=?,pending_generation=NULL,pending_full=0,
                    pending_url=NULL,cursor=?,seed_done=?,synced_at=?,complete=1,error=NULL WHERE backend=? AND id=?""", (name, generation, cursor, int(bool(seed_done if seed_done is not None else state["seed_done"])), _now(), backend, folder_id))
            else:
                self.db.execute("UPDATE folders SET name=?,pending_url=?,seed_done=?,error=NULL WHERE backend=? AND id=?", (name, next_url, int(bool(seed_done if seed_done is not None else state["seed_done"])), backend, folder_id))
            if snapshot_mode is not None:
                self.db.execute("UPDATE folders SET snapshot_mode=? WHERE backend=? AND id=?", (int(snapshot_mode), backend, folder_id))
        self._private_files()

    def failed(self, backend: str, folder_id: str, name: str, error: str):
        with self.db:
            self.db.execute("UPDATE folders SET name=?,error=?,complete=0 WHERE backend=? AND id=?", (name, error, backend, folder_id))

    def set_inventory(self, backend: str, live_ids):
        with self.db:
            self.db.execute("INSERT OR REPLACE INTO meta VALUES(?,?)", ("inventory:" + backend, _dump(sorted(set(live_ids)))))

    def reconcile_folders(self, backend: str, live_ids: set[str]):
        """Only invoke after fully enumerating the current folder hierarchy."""
        with self.db:
            self.db.execute("INSERT OR REPLACE INTO meta VALUES(?,?)", ("inventory:" + backend, _dump(sorted(live_ids))))
            stale = [row[0] for row in self.db.execute("SELECT id FROM folders WHERE backend=?", (backend,)) if row[0] not in live_ids]
            for identity in stale:
                self._delete_where("backend=? AND folder_id=?", (backend, identity))
                self.db.execute("DELETE FROM folders WHERE backend=? AND id=?", (backend, identity))

    def set_cooldown(self, seconds: float):
        with self.db:
            self.db.execute("INSERT OR REPLACE INTO meta VALUES('cooldown_until',?)", (str(time.time() + max(0, seconds)),))

    def cooldown_remaining(self) -> float:
        row = self.db.execute("SELECT value FROM meta WHERE key='cooldown_until'").fetchone()
        return max(0, float(row[0]) - time.time()) if row else 0

    def status(self, backend="rest", folder=None, *, include_stats=True) -> dict:
        rows = [dict(row) for row in self.db.execute("SELECT * FROM folders WHERE backend=? ORDER BY name", (backend,))]
        inventory = self.db.execute("SELECT value FROM meta WHERE key=?", ("inventory:" + backend,)).fetchone()
        live = set(json.loads(inventory[0])) if inventory else None
        whole = live is not None and {row["id"] for row in rows if row["complete"] and not row["pending_generation"]} == live
        if folder:
            rows = [row for row in rows if row["id"] == folder or fold(row["name"]) == fold(folder)]
        for row in rows:
            # Delta/next links are opaque provider state; do not emit their URLs.
            row["resumable"] = bool(row.pop("pending_generation"))
            row.pop("pending_url")
            row.pop("cursor")
            row.pop("active_generation")
            row.pop("options_signature")
        result = {"backend": backend, "local": True, "path": str(self.path), "schema_version": SCHEMA_VERSION,
                "scope": "whole_mailbox" if whole and not folder else "explicitly_synced_folders", "whole_mailbox_complete": whole,
                "complete": bool(rows) and all(row["complete"] and not row["resumable"] for row in rows), "folders": rows,
                "folder_count": len(rows), "discovered_folder_count": len(live) if live is not None else None,
                "oldest_sync": min((row["synced_at"] for row in rows if row["synced_at"]), default=None),
                "freshness": "as_of_folder_sync; run local sync to refresh",
                "cooldown_seconds": round(self.cooldown_remaining(), 3)}
        if include_stats:
            count = self.db.execute("""SELECT count(*) FROM messages m JOIN folders f ON f.backend=m.backend AND f.id=m.folder_id AND f.active_generation=m.generation WHERE m.backend=?""", (backend,)).fetchone()[0]
            files = {suffix or "database": Path(str(self.path) + suffix).stat().st_size for suffix in ("", "-wal", "-shm") if Path(str(self.path) + suffix).exists()}
            body_stats = self.db.execute("SELECT count(*),coalesce(sum(original_bytes),0),coalesce(sum(text_bytes),0),coalesce(sum(length(original)+coalesce(length(plain),0)),0) FROM bodies").fetchone()
            result.update(message_count=count, file_bytes=files, database_bytes=sum(files.values()),
                          bodies={"unique": body_stats[0], "original_bytes": body_stats[1], "text_bytes": body_stats[2], "compressed_bytes": body_stats[3]})
        else:
            result.pop("folders")
            result.pop("path")
            result.pop("cooldown_seconds")
        return result

    def query(self, query="", *, backend="rest", limit=25, folder=None, sender=None, recipient=None,
              domain=None, person=None, after=None, before=None, conversation=None, has_attachments=None,
              require_complete=False, match_mode="prefix", snippet_chars=240, offset=0,
              exclude_id=None) -> tuple[list[dict], dict]:
        if not 1 <= limit <= 100000 or offset < 0:
            raise ValueError("limit must be 1..100000 and offset nonnegative")
        if not 40 <= snippet_chars <= 2000:
            raise ValueError("snippet_chars must be between 40 and 2000")
        terms, fields = _terms(query)
        values = {"folder": folder, "sender": sender, "recipient": recipient, "domain": domain, "person": person,
                  "after": after, "before": before, "conversation": conversation, "has_attachments": has_attachments}
        values = {key: value if value is not None else fields.get(key) for key, value in values.items()}
        match, highlights = _fts_query(terms, match_mode)
        self.db.execute("BEGIN")
        try:
            meta = self.status(backend, values["folder"], include_stats=False)
            if require_complete and not meta["complete"]:
                raise OutlookCliError("The selected local scope is incomplete. Run local sync successfully first.")
            clauses, args = ["m.backend=?"], [backend]
            if exclude_id is not None:
                clauses.append("m.id<>?")
                args.append(exclude_id)
            join = "JOIN folders f ON f.backend=m.backend AND f.id=m.folder_id AND f.active_generation=m.generation"
            if match:
                join += " JOIN mail_fts ON mail_fts.rowid=m.record_id"
                clauses.append("mail_fts MATCH ?")
                args.append(match)
            if values["folder"]:
                selected = [row[0] for row in self.db.execute("SELECT id,name FROM folders WHERE backend=?", (backend,)) if row[0] == values["folder"] or fold(row[1]) == fold(values["folder"])]
                clauses.append("m.folder_id IN (" + ",".join("?" for _ in selected) + ")" if selected else "0")
                args.extend(selected)
            if values["conversation"]:
                clauses.append("m.conversation_id=?")
                args.append(values["conversation"])
            if values["has_attachments"] is not None:
                clauses.append("m.has_attachments=?")
                args.append(int(bool(values["has_attachments"])))
            for name, operator in (("after", ">="), ("before", "<")):
                if values[name]:
                    clauses.append(f"m.received {operator} ?")
                    args.append(_date(values[name]))
            for name, role in (("sender", "a.role='from' AND "), ("recipient", "a.role IN ('to','cc','bcc') AND "), ("domain", ""), ("person", "")):
                if not values[name]:
                    continue
                value = fold(values[name]).lstrip("@") if name == "domain" else fold(values[name])
                if name == "domain":
                    condition = "a.domain=?"
                elif name == "person":
                    condition = "(a.address=? OR instr(a.name,?)>0)"
                else:
                    condition = "a.address=?"
                clauses.append(f"EXISTS(SELECT 1 FROM addresses a WHERE a.message_rowid=m.record_id AND {role}{condition})")
                args.extend([value, value] if name == "person" else [value])
            where = " AND ".join(clauses)
            total = self.db.execute(f"SELECT count(*) FROM messages m {join} WHERE {where}", args).fetchone()[0]
            score = "bm25(mail_fts,8.0,3.0,1.0,2.0)" if match else "0.0"
            order = "score ASC,m.received DESC,m.id" if match else "m.received DESC,m.id"
            rows = self.db.execute(f"SELECT m.*,f.name AS folder_name,{score} AS score FROM messages m {join} WHERE {where} ORDER BY {order} LIMIT ? OFFSET ?", args + [limit, offset]).fetchall()
            data = []
            for row in rows:
                record = _unpack(row["metadata"])
                plain = self._body(row["body_hash"]) or record.get("preview", "")
                source = plain
                if highlights and not all(term in fold(plain) for term in highlights):
                    people = [record.get("sender") or {}] + record.get("to", []) + record.get("cc", [])
                    source = " ".join([record.get("subject", ""),
                                       " ".join(person.get("name", "") + " " + person.get("address", "") for person in people),
                                       " ".join(item.get("name", "") for item in record.get("attachments", [])), plain])
                snippet, highlighted = _snippet(source, highlights, snippet_chars)
                data.append({key: record.get(key) for key in ("id", "subject", "sender", "to", "cc", "received", "conversation_id", "is_read", "categories")})
                data[-1].update(backend=backend, folder_id=row["folder_id"], folder_name=row["folder_name"], has_attachments=bool(row["has_attachments"]), snippet=snippet, highlighted=highlighted, score=row["score"], attachment_count=len(record.get("attachments", [])))
            meta.update(total_matches=total, returned_count=len(data), offset=offset, has_more=total > offset + len(data),
                        result_complete=meta["complete"] and total <= offset + len(data), match_mode=match_mode,
                        body_scope="stored bodies where available, otherwise previews; attachment metadata only")
            return data, meta
        finally:
            self.db.rollback()

    def _find(self, identity: str, backend: str):
        row = self.db.execute("""SELECT m.*,f.name AS folder_name FROM messages m JOIN folders f ON
            f.backend=m.backend AND f.id=m.folder_id AND f.active_generation=m.generation
            WHERE m.backend=? AND m.id=? ORDER BY m.received DESC LIMIT 1""", (backend, identity)).fetchone()
        if row is None:
            raise ResourceNotFoundError("Message not found in the local store; run local sync to populate it")
        return row

    def _read_row(self, row, backend, body_format):
        if body_format not in ("text", "html", "none", "preview"):
            raise ValueError("body_format must be text, html, none or preview")
        record = _unpack(row["metadata"])
        record.update(backend=backend, folder_id=row["folder_id"], folder_name=row["folder_name"], local=True)
        if body_format != "none":
            record["body"] = record.get("preview", "") if body_format == "preview" else self._body(row["body_hash"], original=body_format == "html")
            if body_format != "html":
                record["body_type"] = "Text"
        else:
            record.pop("body_type", None)
        return record

    def read(self, identity: str, *, backend="rest", body_format="text") -> dict:
        self.db.execute("BEGIN")
        try:
            return self._read_row(self._find(identity, backend), backend, body_format)
        finally:
            self.db.rollback()

    get = read

    def thread(self, identity: str, *, backend="rest", limit=1000, body_format="text") -> tuple[list[dict], dict]:
        if not 1 <= limit <= 100000:
            raise ValueError("limit must be 1..100000")
        self.db.execute("BEGIN")
        try:
            try:
                row = self._find(identity, backend)
                conversation = row["conversation_id"]
                if not conversation:
                    return [self._read_row(row, backend, body_format)], {"local": True, "returned_count": 1, "conversation_id": "", "complete": False, "reason": "missing_conversation_id"}
            except ResourceNotFoundError:
                conversation = identity
            rows = self.db.execute("""SELECT m.*,f.name AS folder_name FROM messages m JOIN folders f ON f.backend=m.backend AND f.id=m.folder_id
                AND f.active_generation=m.generation WHERE m.backend=? AND m.conversation_id=? ORDER BY m.received,m.id LIMIT ?""", (backend, conversation, limit + 1)).fetchall()
            if not rows:
                raise ResourceNotFoundError("Conversation not found in the local store")
            meta = self.status(backend, include_stats=False)
            meta.update(conversation_id=conversation, returned_count=min(len(rows), limit), has_more=len(rows) > limit,
                        result_complete=meta["whole_mailbox_complete"] and len(rows) <= limit)
            return [self._read_row(item, backend, body_format) for item in rows[:limit]], meta
        finally:
            self.db.rollback()

    def related(self, identity: str, *, backend="rest", limit=25) -> tuple[list[dict], dict]:
        record = self.read(identity, backend=backend, body_format="none")
        address = (record.get("sender") or {}).get("address")
        if not address:
            return [], {"local": True, "returned_count": 0, "reason": "missing_sender"}
        data, meta = self.query(backend=backend, person=address, limit=limit, exclude_id=identity)
        meta.update(returned_count=len(data), related_person=address)
        return data, meta

    def attachments(self, identity: str, *, backend="rest") -> list[dict]:
        row = self._find(identity, backend)
        return [dict(item) for item in self.db.execute("SELECT id,name,content_type,size,is_inline FROM attachments WHERE message_rowid=? ORDER BY name,id", (row["record_id"],))]

    def compact(self) -> dict:
        """Explicit maintenance, never erase the retained legacy index."""
        with self.db:
            self.db.execute("DELETE FROM bodies WHERE hash NOT IN (SELECT body_hash FROM messages WHERE body_hash IS NOT NULL)")
            self.db.execute("INSERT INTO mail_fts(mail_fts) VALUES('optimize')")
        self.db.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        self.db.execute("VACUUM")
        self._private_files()
        return self.status()

    def migrate_legacy(self, source: Path, *, batch_size=100) -> dict:
        source = Path(source)
        if source.resolve() == self.path.resolve():
            raise ValueError("Legacy source and new mail store must differ")
        legacy = sqlite3.connect(source.resolve().as_uri() + "?mode=ro", uri=True)
        legacy.row_factory = sqlite3.Row
        legacy.execute("PRAGMA query_only=ON")
        migrated = skipped = 0
        try:
            legacy.execute("BEGIN")
            for folder in legacy.execute("SELECT * FROM folders ORDER BY backend,id").fetchall():
                backend, identity, name = folder["backend"], folder["id"], folder["name"]
                if self.folder_state(backend, identity).get("active_generation"):
                    skipped += 1
                    continue
                signature = f"legacy:{source.stat().st_size}:{source.stat().st_mtime_ns}"
                self.begin_sync(backend, identity, name, full=True, include_body=bool(folder["include_body"]), options_signature=signature)
                rows = legacy.execute("SELECT payload FROM messages WHERE backend=? AND folder_id=? ORDER BY rowid", (backend, identity))
                while True:
                    batch = rows.fetchmany(batch_size)
                    if not batch:
                        break
                    self.apply_page(backend, identity, name, [json.loads(item[0]) for item in batch])
                    migrated += len(batch)
                self.apply_page(backend, identity, name, [], complete=True, cursor=folder["cursor"])
                # Preserve source completeness/freshness; no mailbox delta proof is invented.
                with self.db:
                    self.db.execute("UPDATE folders SET synced_at=?,complete=?,error=? WHERE backend=? AND id=?", (folder["synced_at"], folder["complete"], folder["error"], backend, identity))
        finally:
            legacy.close()
        return {"local": True, "source_retained": True, "source": str(source), "migrated_messages": migrated, "skipped_folders": skipped, **self.status()}

    def purge(self):
        self.close()
        for suffix in ("", "-wal", "-shm"):
            path = Path(str(self.path) + suffix)
            if path.is_symlink():
                raise OutlookCliError("Refusing to remove a symlink at the local store path")
            path.unlink(missing_ok=True)
