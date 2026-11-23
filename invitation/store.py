"""SQLite storage for invitation events, their responses and the rate-limit hits.

Its own database file (never users.db). Opened lazily, and the schema is only
CREATE ... IF NOT EXISTS, so opening it on the live server is additive and safe.
"""
import sqlite3
import threading
import time
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    slug TEXT NOT NULL UNIQUE,
    status TEXT NOT NULL DEFAULT 'open' CHECK (status IN ('open', 'closed')),
    closes_at REAL,                 -- optional automatic close (unix time); NULL = none
    image_key TEXT NOT NULL,        -- file name under the invitation files dir
    image_mime TEXT NOT NULL,
    image_width INTEGER NOT NULL,
    image_height INTEGER NOT NULL,
    created_by TEXT NOT NULL,
    created_at REAL NOT NULL,
    updated_by TEXT,
    updated_at REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS responses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id INTEGER NOT NULL REFERENCES events(id) ON DELETE CASCADE,
    full_name TEXT NOT NULL,
    email TEXT NOT NULL,
    email_normalized TEXT NOT NULL,
    attending INTEGER NOT NULL CHECK (attending IN (0, 1)),
    has_allergy INTEGER,
    allergy_notes TEXT,
    submission_count INTEGER NOT NULL DEFAULT 1,
    created_at REAL NOT NULL,
    updated_at REAL NOT NULL,
    UNIQUE (event_id, email_normalized),
    CHECK ((attending = 0 AND has_allergy IS NULL AND allergy_notes IS NULL)
        OR (attending = 1 AND has_allergy = 0 AND allergy_notes IS NULL)
        OR (attending = 1 AND has_allergy = 1 AND allergy_notes IS NOT NULL
            AND length(allergy_notes) BETWEEN 1 AND 200))
);
CREATE TABLE IF NOT EXISTS hits (
    key TEXT NOT NULL,
    at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS responses_event ON responses(event_id, updated_at);
CREATE INDEX IF NOT EXISTS hits_key ON hits(key, at);
"""

FILTERS = {
    "all": "",
    "attending": " AND attending=1",
    "not-attending": " AND attending=0",
    "allergy": " AND has_allergy=1",
}


def is_open(event, now=None):
    """The one rule for 'registration is open': manual switch AND auto-close time."""
    if event is None or event["status"] != "open":
        return False
    closes = event.get("closes_at")
    return closes is None or (now or time.time()) < closes


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
        con.execute("PRAGMA foreign_keys=ON")
        return con

    # -------------------------------------------------------------- events
    def create_event(self, name, slug, closes_at, image, user):
        now = time.time()
        with self._con() as con:
            cur = con.execute(
                "INSERT INTO events (name, slug, status, closes_at, image_key, image_mime,"
                " image_width, image_height, created_by, created_at, updated_by, updated_at)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                (name, slug, "open", closes_at, image["key"], image["mime"], image["width"],
                 image["height"], user, now, user, now))
            return cur.lastrowid

    def update_event(self, eid, user, **fields):
        allowed = {"name", "slug", "status", "closes_at", "image_key", "image_mime",
                   "image_width", "image_height"}
        assert set(fields) <= allowed, fields
        if not fields:
            return
        cols = ", ".join(f"{k}=?" for k in fields)
        with self._con() as con:
            con.execute(f"UPDATE events SET {cols}, updated_by=?, updated_at=? WHERE id=?",
                        (*fields.values(), user, time.time(), eid))

    def delete_event(self, eid):
        with self._con() as con:
            con.execute("DELETE FROM responses WHERE event_id=?", (eid,))
            con.execute("DELETE FROM events WHERE id=?", (eid,))

    def event(self, eid):
        with self._con() as con:
            row = con.execute("SELECT * FROM events WHERE id=?", (eid,)).fetchone()
        return dict(row) if row else None

    def event_by_slug(self, slug):
        with self._con() as con:
            row = con.execute("SELECT * FROM events WHERE slug=?", (slug,)).fetchone()
        return dict(row) if row else None

    def slug_taken(self, slug, except_id=None):
        with self._con() as con:
            row = con.execute("SELECT id FROM events WHERE slug=? AND id IS NOT ?",
                              (slug, except_id)).fetchone()
        return row is not None

    def events(self):
        with self._con() as con:
            rows = con.execute(
                "SELECT e.*, COUNT(r.id) AS total, COALESCE(SUM(r.attending),0) AS attending,"
                " COALESCE(SUM(r.has_allergy),0) AS allergies, MAX(r.updated_at) AS last_response"
                " FROM events e LEFT JOIN responses r ON r.event_id=e.id"
                " GROUP BY e.id ORDER BY e.created_at DESC").fetchall()
        return [dict(r) for r in rows]

    # ----------------------------------------------------------- responses
    def upsert_response(self, eid, value):
        """Insert, or update the row with the same (event, lower(trim(email)))."""
        now = time.time()
        with self._con() as con:
            con.execute(
                "INSERT INTO responses (event_id, full_name, email, email_normalized, attending,"
                " has_allergy, allergy_notes, submission_count, created_at, updated_at)"
                " VALUES (?,?,?,?,?,?,?,1,?,?)"
                " ON CONFLICT (event_id, email_normalized) DO UPDATE SET"
                " full_name=excluded.full_name, email=excluded.email, attending=excluded.attending,"
                " has_allergy=excluded.has_allergy, allergy_notes=excluded.allergy_notes,"
                " submission_count=responses.submission_count+1, updated_at=excluded.updated_at",
                (eid, value["full_name"], value["email"], value["email_normalized"],
                 int(value["attending"]),
                 None if value["has_allergy"] is None else int(value["has_allergy"]),
                 value["allergy_notes"], now, now))

    def counts(self, eid):
        with self._con() as con:
            r = con.execute(
                "SELECT COUNT(*) AS total, COALESCE(SUM(attending),0) AS attending,"
                " COALESCE(SUM(1-attending),0) AS not_attending,"
                " COALESCE(SUM(has_allergy),0) AS allergies FROM responses WHERE event_id=?",
                (eid,)).fetchone()
        return dict(r)

    def responses(self, eid, q="", filt="all"):
        sql = "SELECT * FROM responses WHERE event_id=?" + FILTERS.get(filt, "")
        args = [eid]
        q = (q or "").strip().lower()
        if q:
            like = "%" + q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
            sql += " AND (lower(full_name) LIKE ? ESCAPE '\\' OR email_normalized LIKE ? ESCAPE '\\')"
            args += [like, like]
        sql += " ORDER BY updated_at DESC, id DESC"
        with self._con() as con:
            rows = con.execute(sql, args).fetchall()
        return [dict(r) for r in rows]

    def response(self, eid, rid):
        with self._con() as con:
            row = con.execute("SELECT * FROM responses WHERE id=? AND event_id=?",
                              (rid, eid)).fetchone()
        return dict(row) if row else None

    def delete_response(self, eid, rid):
        with self._con() as con:
            return con.execute("DELETE FROM responses WHERE id=? AND event_id=?",
                               (rid, eid)).rowcount

    # ---------------------------------------------------------- rate limit
    def hits(self, key, window):
        """How many times `key` was recorded in the last `window` seconds.
        Kept in SQLite so all gunicorn workers share the count."""
        now = time.time()
        with self._con() as con:
            con.execute("DELETE FROM hits WHERE at<?", (now - 86400,))
            return con.execute("SELECT COUNT(*) FROM hits WHERE key=? AND at>?",
                               (key, now - window)).fetchone()[0]

    def hit(self, key):
        with self._con() as con:
            con.execute("INSERT INTO hits (key, at) VALUES (?,?)", (key, time.time()))
