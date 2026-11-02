"""SQLite storage for partner requests, signing links and the audit log.

Its own database file (never users.db), opened lazily so importing the app
does not create anything until the feature is used.
"""
import datetime as dt
import json
import secrets
import sqlite3
import threading
import time
from pathlib import Path

WIB = dt.timezone(dt.timedelta(hours=7), "WIB")

STATUS = {
    "pending": "Pending",
    "approved": "Approved (Awaiting Signature)",
    "signed": "Signed \u2013 Under Review",
    "completed": "Completed",
    "rejected": "Rejected",
    "resign": "Re-sign requested",
}
SIGNABLE = ("approved", "resign")

SCHEMA = """
CREATE TABLE IF NOT EXISTS requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    request_type TEXT NOT NULL,
    status TEXT NOT NULL,
    data TEXT NOT NULL,
    partner_name TEXT,
    partner_email TEXT,
    role TEXT,
    created_by TEXT NOT NULL,
    created_at REAL NOT NULL,
    updated_at REAL NOT NULL,
    reject_reason TEXT,
    resign_note TEXT,
    agreement_note TEXT,      -- why there is no link (template missing, PDF failed)
    email_note TEXT,          -- result of the last e-mail attempt
    agreement_pdf TEXT,       -- relative to the files dir
    agreement_docx TEXT,
    signed_file TEXT,         -- signed PDF or the partner's upload
    signed_docx TEXT,
    signed_kind TEXT,         -- drawn | typed | upload
    signed_at REAL,
    signer_ip TEXT,
    signer_ua TEXT
);
CREATE TABLE IF NOT EXISTS tokens (
    token TEXT PRIMARY KEY,
    request_id INTEGER NOT NULL REFERENCES requests(id),
    created_at REAL NOT NULL,
    expires_at REAL NOT NULL,
    used_at REAL,
    revoked_at REAL
);
CREATE TABLE IF NOT EXISTS audit (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    request_id INTEGER NOT NULL REFERENCES requests(id),
    at REAL NOT NULL,
    actor TEXT NOT NULL,
    from_status TEXT,
    to_status TEXT,
    note TEXT
);
CREATE INDEX IF NOT EXISTS tokens_request ON tokens(request_id);
CREATE INDEX IF NOT EXISTS audit_request ON audit(request_id);
"""


def wib(ts, fmt="%Y-%m-%d %H:%M WIB"):
    if not ts:
        return ""
    return dt.datetime.fromtimestamp(ts, WIB).strftime(fmt)


def today_wib():
    return dt.datetime.now(WIB).date()


def ref(request_id):
    return f"PA-{int(request_id):05d}"


class Store:
    def __init__(self, path):
        self.path = Path(path)
        self._ready = False
        self._lock = threading.Lock()

    def _con(self):
        if not self._ready:
            with self._lock:
                if not self._ready:
                    self.path.parent.mkdir(parents=True, exist_ok=True)
                    with sqlite3.connect(self.path) as con:
                        con.execute("PRAGMA journal_mode=WAL")
                        con.executescript(SCHEMA)
                    self._ready = True
        con = sqlite3.connect(self.path, timeout=15)
        con.row_factory = sqlite3.Row
        return con

    # ------------------------------------------------------------ requests
    def create(self, data, user):
        now = time.time()
        with self._con() as con:
            cur = con.execute(
                "INSERT INTO requests (request_type, status, data, partner_name, partner_email,"
                " role, created_by, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
                (data["request_type"], "pending", json.dumps(data, ensure_ascii=False),
                 data.get("partner_name"), data.get("partner_email"), data.get("role"),
                 user, now, now))
            rid = cur.lastrowid
            con.execute("INSERT INTO audit (request_id, at, actor, from_status, to_status, note)"
                        " VALUES (?,?,?,?,?,?)", (rid, now, user, None, "pending", "Submitted"))
        return rid

    def get(self, rid):
        with self._con() as con:
            row = con.execute("SELECT * FROM requests WHERE id=?", (rid,)).fetchone()
        return _req(row)

    def list(self, user=None, status=None):
        sql, args = "SELECT * FROM requests WHERE 1=1", []
        if user is not None:
            sql += " AND created_by=?"
            args.append(user)
        if status:
            sql += " AND status=?"
            args.append(status)
        with self._con() as con:
            rows = con.execute(sql + " ORDER BY id DESC", args).fetchall()
        return [_req(r) for r in rows]

    def counts(self):
        with self._con() as con:
            rows = con.execute("SELECT status, COUNT(*) n FROM requests GROUP BY status").fetchall()
        return {r["status"]: r["n"] for r in rows}

    def update(self, rid, **fields):
        if not fields:
            return
        fields["updated_at"] = time.time()
        cols = ", ".join(f"{k}=?" for k in fields)
        with self._con() as con:
            con.execute(f"UPDATE requests SET {cols} WHERE id=?", (*fields.values(), rid))

    def set_status(self, rid, new, actor, note=None, expect=None, **fields):
        """Change status and write the audit row in one transaction.

        `expect` (tuple of statuses) makes the change conditional, so two admins
        or a double click cannot both move the same request. Returns True when
        the status changed.
        """
        now = time.time()
        with self._con() as con:
            row = con.execute("SELECT status FROM requests WHERE id=?", (rid,)).fetchone()
            if row is None or (expect and row["status"] not in expect):
                return False
            fields.update(status=new, updated_at=now)
            cols = ", ".join(f"{k}=?" for k in fields)
            con.execute(f"UPDATE requests SET {cols} WHERE id=?", (*fields.values(), rid))
            con.execute("INSERT INTO audit (request_id, at, actor, from_status, to_status, note)"
                        " VALUES (?,?,?,?,?,?)", (rid, now, actor, row["status"], new, note))
        return True

    def log(self, rid, actor, note):
        """Audit row without a status change (e-mail sent, link created, ...)."""
        with self._con() as con:
            st = con.execute("SELECT status FROM requests WHERE id=?", (rid,)).fetchone()
            s = st["status"] if st else None
            con.execute("INSERT INTO audit (request_id, at, actor, from_status, to_status, note)"
                        " VALUES (?,?,?,?,?,?)", (rid, time.time(), actor, s, s, note))

    def audit(self, rid):
        with self._con() as con:
            return con.execute("SELECT * FROM audit WHERE request_id=? ORDER BY id", (rid,)).fetchall()

    # -------------------------------------------------------------- tokens
    def new_token(self, rid, days):
        """Revoke every earlier link of the request and issue a fresh one."""
        token = secrets.token_urlsafe(32)
        now = time.time()
        with self._con() as con:
            con.execute("UPDATE tokens SET revoked_at=? WHERE request_id=? AND revoked_at IS NULL"
                        " AND used_at IS NULL", (now, rid))
            con.execute("INSERT INTO tokens (token, request_id, created_at, expires_at)"
                        " VALUES (?,?,?,?)", (token, rid, now, now + days * 86400))
        return token

    def revoke(self, rid):
        """Close every open link of the request (rejected)."""
        with self._con() as con:
            con.execute("UPDATE tokens SET revoked_at=? WHERE request_id=? AND revoked_at IS NULL"
                        " AND used_at IS NULL", (time.time(), rid))

    def token(self, token):
        if not token or len(token) > 200:
            return None
        with self._con() as con:
            return con.execute("SELECT * FROM tokens WHERE token=?", (token,)).fetchone()

    def current_token(self, rid):
        with self._con() as con:
            return con.execute("SELECT * FROM tokens WHERE request_id=? AND revoked_at IS NULL"
                               " ORDER BY created_at DESC LIMIT 1", (rid,)).fetchone()

    def use_token(self, token):
        """Mark a link used. False when it was already used/revoked (race-safe)."""
        with self._con() as con:
            cur = con.execute("UPDATE tokens SET used_at=? WHERE token=? AND used_at IS NULL"
                              " AND revoked_at IS NULL AND expires_at>?",
                              (time.time(), token, time.time()))
            return cur.rowcount == 1


def _req(row):
    if row is None:
        return None
    r = dict(row)
    r["data"] = json.loads(r["data"])
    r["ref"] = ref(r["id"])
    r["status_label"] = STATUS.get(r["status"], r["status"])
    return r
