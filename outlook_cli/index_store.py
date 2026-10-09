"""Account-local FTS index with transactional folder snapshots and delta cursors."""
from __future__ import annotations

import json
import os
import shlex
import sqlite3
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

from .exceptions import OutlookCliError


def _now():
    return datetime.now(timezone.utc).isoformat()


def _text(value: str) -> str:
    return unicodedata.normalize("NFKD", value.casefold()).replace("ı", "i")


def _encode(value):
    return json.dumps(value, ensure_ascii=False, default=lambda x: x.isoformat())


class MailIndex:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.db = sqlite3.connect(self.path, timeout=30)
        os.chmod(self.path, 0o600)
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.executescript("""
        CREATE TABLE IF NOT EXISTS messages (
          backend TEXT NOT NULL, id TEXT NOT NULL, folder_id TEXT NOT NULL,
          received TEXT NOT NULL, conversation_id TEXT NOT NULL, payload TEXT NOT NULL,
          search_text TEXT NOT NULL, UNIQUE(backend,id));
        CREATE INDEX IF NOT EXISTS messages_scope ON messages(backend,folder_id,received);
        CREATE INDEX IF NOT EXISTS messages_conversation ON messages(backend,conversation_id);
        CREATE VIRTUAL TABLE IF NOT EXISTS message_fts USING fts5(search_text, content='messages', content_rowid='rowid', tokenize='unicode61 remove_diacritics 2');
        CREATE TRIGGER IF NOT EXISTS message_insert AFTER INSERT ON messages BEGIN
          INSERT INTO message_fts(rowid,search_text) VALUES(new.rowid,new.search_text); END;
        CREATE TRIGGER IF NOT EXISTS message_delete AFTER DELETE ON messages BEGIN
          INSERT INTO message_fts(message_fts,rowid,search_text) VALUES('delete',old.rowid,old.search_text); END;
        CREATE TRIGGER IF NOT EXISTS message_update AFTER UPDATE ON messages BEGIN
          INSERT INTO message_fts(message_fts,rowid,search_text) VALUES('delete',old.rowid,old.search_text);
          INSERT INTO message_fts(rowid,search_text) VALUES(new.rowid,new.search_text); END;
        CREATE TABLE IF NOT EXISTS folders (
          backend TEXT NOT NULL, id TEXT NOT NULL, name TEXT NOT NULL, synced_at TEXT,
          complete INTEGER NOT NULL DEFAULT 0, cursor TEXT, include_body INTEGER NOT NULL DEFAULT 0,
          error TEXT, PRIMARY KEY(backend,id));
        CREATE TABLE IF NOT EXISTS addresses (
          backend TEXT NOT NULL,id TEXT NOT NULL,address TEXT NOT NULL,role TEXT NOT NULL,
          PRIMARY KEY(backend,id,address,role),
          FOREIGN KEY(backend,id) REFERENCES messages(backend,id) ON DELETE CASCADE);
        CREATE INDEX IF NOT EXISTS address_lookup ON addresses(backend,address,role);
        """)

    def close(self):
        self.db.close()

    def folder_state(self, backend: str, folder_id: str) -> dict:
        row = self.db.execute("SELECT * FROM folders WHERE backend=? AND id=?", (backend, folder_id)).fetchone()
        return dict(row) if row else {}

    def _put(self, backend, folder_id, message):
        message = dict(message)
        message.pop("display_num", None)
        text = " ".join(str(message.get(k) or "") for k in ("subject", "preview", "body"))
        people = [(message.get("sender") or {}, "from")]
        people += [(x, "to") for x in message.get("to", [])] + [(x, "cc") for x in message.get("cc", [])]
        text += " " + " ".join(f"{p.get('name','')} {p.get('address','')}" for p, _ in people)
        received = message.get("received") or ""
        if hasattr(received, "isoformat"):
            received = received.isoformat()
        if received:
            parsed = datetime.fromisoformat(received.replace("Z", "+00:00"))
            received = parsed.replace(tzinfo=parsed.tzinfo or timezone.utc).astimezone(timezone.utc).isoformat()
        self.db.execute("""INSERT INTO messages(backend,id,folder_id,received,conversation_id,payload,search_text)
            VALUES(?,?,?,?,?,?,?) ON CONFLICT(backend,id) DO UPDATE SET folder_id=excluded.folder_id,
            received=excluded.received,conversation_id=excluded.conversation_id,payload=excluded.payload,search_text=excluded.search_text""",
            (backend,message["id"],folder_id,received,message.get("conversation_id") or "",_encode(message),_text(text)))
        self.db.execute("DELETE FROM addresses WHERE backend=? AND id=?", (backend,message["id"]))
        self.db.executemany("INSERT OR IGNORE INTO addresses VALUES(?,?,?,?)",
            [(backend,message["id"],p["address"].casefold(),role) for p,role in people if p.get("address")])

    def apply(self, backend: str, folder_id: str, name: str, messages: list[dict], *,
              full: bool, cursor: str | None = None, removed=(), include_body=False):
        """Commit data AND checkpoint together, only after a complete fetch."""
        with self.db:
            if full:
                self.db.execute("DELETE FROM messages WHERE backend=? AND folder_id=?", (backend,folder_id))
            for identity in removed:
                # A tombstone from the old folder must not remove an item already moved elsewhere.
                self.db.execute("DELETE FROM messages WHERE backend=? AND folder_id=? AND id=?", (backend,folder_id,identity))
            for message in messages:
                self._put(backend,folder_id,message)
            self.db.execute("""INSERT INTO folders VALUES(?,?,?,?,?,?,?,NULL)
                ON CONFLICT(backend,id) DO UPDATE SET name=excluded.name,synced_at=excluded.synced_at,
                complete=1,cursor=excluded.cursor,include_body=excluded.include_body,error=NULL""",
                (backend,folder_id,name,_now(),1,cursor,int(include_body)))

    def failed(self, backend: str, folder_id: str, name: str, error: str):
        with self.db:
            self.db.execute("""INSERT INTO folders(backend,id,name,complete,error) VALUES(?,?,?,0,?)
                ON CONFLICT(backend,id) DO UPDATE SET complete=0,error=excluded.error""", (backend,folder_id,name,error))

    def reconcile_folders(self, backend: str, live_ids: set[str]):
        """Only call after a complete hierarchy scan to remove deleted folders."""
        with self.db:
            stale = [r[0] for r in self.db.execute("SELECT id FROM folders WHERE backend=?", (backend,)) if r[0] not in live_ids]
            for identity in stale:
                self.db.execute("DELETE FROM messages WHERE backend=? AND folder_id=?", (backend,identity))
                self.db.execute("DELETE FROM folders WHERE backend=? AND id=?", (backend,identity))

    def status(self, backend: str = "rest", folder: str | None = None) -> dict:
        args = [backend]
        where = "backend=?"
        if folder:
            where += " AND (id=? OR name=?)"
            args += [folder,folder]
        rows = [dict(row) for row in self.db.execute(f"SELECT backend,id,name,synced_at,complete,include_body,error FROM folders WHERE {where} ORDER BY name",args)]
        return {"backend":backend,"scope":"explicitly_synced_folders", "complete":bool(rows) and all(r["complete"] for r in rows),
                "folders":rows,"oldest_sync":min((r["synced_at"] for r in rows if r["synced_at"]),default=None),
                "local":True,"freshness":"as_of_folder_sync; run index sync to refresh"}

    def query(self, query: str = "", *, backend="rest", limit=25, folder=None, sender=None,
              recipient=None, domain=None, after=None, before=None, conversation=None,
              require_complete=False) -> tuple[list[dict],dict]:
        # One WAL read snapshot for scope, totals, and rows even during a sync.
        self.db.execute("BEGIN")
        try:
            return self._query(query,backend=backend,limit=limit,folder=folder,sender=sender,
                recipient=recipient,domain=domain,after=after,before=before,
                conversation=conversation,require_complete=require_complete)
        finally:
            self.db.rollback()

    def _query(self, query: str = "", *, backend="rest", limit=25, folder=None, sender=None,
               recipient=None, domain=None, after=None, before=None, conversation=None,
               require_complete=False):
        if limit < 1 or limit > 100000:
            raise ValueError("max must be between 1 and 100000")
        meta = self.status(backend,folder)
        if require_complete and not meta["complete"]:
            raise OutlookCliError("The selected index scope is incomplete. Run index sync successfully first.")
        clauses, args = ["m.backend=?"], [backend]
        if query.strip():
            try:
                terms = shlex.split(query)
            except ValueError as exc:
                raise ValueError("Unclosed quotation in local search query") from exc
            literal = " AND ".join('"' + _text(term).replace('"','""') + '"' for term in terms)
            if literal:
                clauses.append("m.rowid IN (SELECT rowid FROM message_fts WHERE message_fts MATCH ?)")
                args.append(literal)
        if folder:
            clauses.append("m.folder_id IN (SELECT id FROM folders WHERE backend=? AND (id=? OR name=?))")
            args += [backend,folder,folder]
        if conversation:
            clauses.append("m.conversation_id=?"); args.append(conversation)
        for value,op in [(after,">="),(before,"<")]:
            if value:
                parsed = datetime.fromisoformat(value.replace("Z","+00:00"))
                if parsed.tzinfo is None:
                    parsed = parsed.replace(tzinfo=timezone.utc)
                clauses.append(f"julianday(m.received) {op} julianday(?)"); args.append(parsed.isoformat())
        if sender:
            clauses.append("EXISTS(SELECT 1 FROM addresses a WHERE a.backend=m.backend AND a.id=m.id AND a.role='from' AND a.address=?)")
            args.append(sender.casefold())
        if recipient:
            clauses.append("EXISTS(SELECT 1 FROM addresses a WHERE a.backend=m.backend AND a.id=m.id AND a.role IN ('to','cc') AND a.address=?)")
            args.append(recipient.casefold())
        if domain:
            clauses.append("EXISTS(SELECT 1 FROM addresses a WHERE a.backend=m.backend AND a.id=m.id AND substr(a.address,instr(a.address,'@')+1)=?)")
            args.append(domain.casefold().lstrip("@"))
        where = " AND ".join(clauses)
        total = self.db.execute(f"SELECT count(*) FROM messages m WHERE {where}",args).fetchone()[0]
        rows = self.db.execute(f"SELECT m.payload,m.folder_id FROM messages m WHERE {where} ORDER BY m.received DESC,m.id LIMIT ?",args+[limit]).fetchall()
        data = [dict(json.loads(row["payload"]),folder_id=row["folder_id"],backend=backend) for row in rows]
        meta.update(total_matches=total,returned_count=len(data),has_more=total>len(data),
                    result_complete=meta["complete"] and total<=len(data),body_scope="body only for folders synced with --include-body")
        return data,meta
