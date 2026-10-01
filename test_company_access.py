#!/usr/bin/env python3
"""Company migration, authorization, navigation, and admin checks."""
import base64
import os
import sqlite3
import tempfile
import time
from pathlib import Path

from werkzeug.security import generate_password_hash

TMP = Path(tempfile.mkdtemp())
DB = TMP / "legacy.db"
admin_hash = generate_password_hash("admin-password-1")
user_hash = generate_password_hash("user-password-1")
with sqlite3.connect(DB) as con:
    con.execute("""CREATE TABLE users (
        username TEXT PRIMARY KEY, pw_hash TEXT NOT NULL,
        is_admin INTEGER NOT NULL DEFAULT 0, created REAL NOT NULL,
        last_seen REAL)""")
    con.execute("INSERT INTO users VALUES (?,?,?,?,?)",
                ("legacy-admin", admin_hash, 1, time.time(), None))
    con.execute("INSERT INTO users VALUES (?,?,?,?,?)",
                ("legacy-user", user_hash, 0, time.time(), None))

os.environ["DW_AUTH_DB"] = str(DB)
os.environ.pop("DW_ADMIN_USER", None)
os.environ.pop("DW_ADMIN_PASS", None)

import auth  # noqa: E402
from app import app  # noqa: E402

app.config["TESTING"] = True


def headers(user, password):
    token = base64.b64encode(f"{user}:{password}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Accept": "text/html"}


ADMIN = headers("legacy-admin", "admin-password-1")
admin = auth.verify("legacy-admin", "admin-password-1")
plain = auth.verify("legacy-user", "user-password-1")
assert (admin["can_dpm"], admin["can_kvb"]) == (1, 1)
assert (plain["can_dpm"], plain["can_kvb"]) == (1, 0)
assert auth.has_company(admin, "dpm") and auth.has_company(admin, "kvb")
assert not auth.has_company(plain, "kvb")
with sqlite3.connect(DB) as con:
    rows = dict(con.execute("SELECT username, pw_hash FROM users"))
assert rows["legacy-admin"] == admin_hash and rows["legacy-user"] == user_hash

c = app.test_client()
# New accounts require at least one company.
r = c.post("/admin/users", headers=ADMIN,
           data={"username": "none", "password": "none-password-1"})
assert r.status_code == 400 and b"at least one company" in r.data
for name, access in (
    ("dpm", {"can_dpm": "1"}),
    ("kvb", {"can_kvb": "1"}),
    ("both", {"can_dpm": "1", "can_kvb": "1"}),
):
    r = c.post("/admin/users", headers=ADMIN,
               data={"username": name, "password": f"{name}-password-1", **access})
    assert r.status_code == 302, (name, r.status_code, r.data[:300])

# Access edits do not touch password hashes.
with sqlite3.connect(DB) as con:
    before = con.execute("SELECT pw_hash FROM users WHERE username='dpm'").fetchone()[0]
r = c.post("/admin/users/dpm/access", headers=ADMIN, data={"can_kvb": "1"})
assert r.status_code == 302
with sqlite3.connect(DB) as con:
    row = con.execute("SELECT pw_hash, can_dpm, can_kvb FROM users WHERE username='dpm'").fetchone()
assert row == (before, 0, 1)
r = c.post("/admin/users/dpm/access", headers=ADMIN, data={})
assert r.status_code == 400 and b"at least one company" in r.data

# Re-running startup migration must not overwrite later access edits.
c.post("/admin/users/legacy-admin/access", headers=ADMIN, data={"can_dpm": "1"})
auth.init_db()
assert auth.verify("legacy-admin", "admin-password-1")["can_kvb"] == 0
c.post("/admin/users/legacy-admin/access", headers=ADMIN,
       data={"can_dpm": "1", "can_kvb": "1"})

clients = {}
for name in ("dpm", "kvb", "both"):
    client = app.test_client()
    client.environ_base["HTTP_AUTHORIZATION"] = headers(name, f"{name}-password-1")["Authorization"]
    client.environ_base["HTTP_ACCEPT"] = "text/html"
    clients[name] = client

assert clients["dpm"].get("/tool/dw").status_code == 403  # edited to KVB-only above
# Restore dpm as DPM-only through the admin endpoint.
c.post("/admin/users/dpm/access", headers=ADMIN, data={"can_dpm": "1"})
assert clients["dpm"].get("/tool/dw").status_code == 200
assert clients["dpm"].get("/kvb/tool/dw").status_code == 403
assert clients["kvb"].get("/kvb/tool/dw").status_code == 200
assert clients["kvb"].get("/tool/dw").status_code == 403
assert clients["kvb"].get("/tool/segregate").status_code == 403
assert clients["both"].get("/tool/segregate").status_code == 200
assert clients["both"].get("/kvb/tool/dw").status_code == 200

# A browser login for a KVB-only user lands in KVB, never on forbidden DPM.
browser = app.test_client()
r = browser.post("/login", data={"username": "kvb", "password": "kvb-password-1"})
assert r.status_code == 302 and r.headers["Location"].endswith("/kvb")

# Navigation exposes both companies but inaccessible one is not a link.
dpm_page = clients["dpm"].get("/").data
assert b"Dupoin DPM Tools" in dpm_page
assert b'href="/tool/dw"' in dpm_page and b'href="/tool/segregate"' in dpm_page
assert b'data-company="kvb" aria-disabled="true"' in dpm_page
kvb_page = clients["kvb"].get("/kvb").data
assert b"KVB Tools" in kvb_page and b"KVB Plus" in kvb_page
assert b"Deal Segregator" not in kvb_page
assert b'data-company="dpm" aria-disabled="true"' in kvb_page

admin_page = c.get("/admin", headers=ADMIN).data
assert b'name="can_dpm"' in admin_page and b'name="can_kvb"' in admin_page
assert b"/admin/users/legacy-user/access" in admin_page

print("OK company access")
