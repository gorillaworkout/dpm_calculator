#!/usr/bin/env python3
"""Company migration, authorization, navigation, and admin checks."""
import base64
import io
import os
import sqlite3
import tempfile
import time
import zipfile
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
from app import JOBS, TOOLS, KVB_TOOLS, _read_state, _write_state, app  # noqa: E402

assert KVB_TOOLS["segregate"][2] == TOOLS["segregate"][2] == ("deal_segregator.py",)
assert "pl-desk" not in KVB_TOOLS and "pl-desk" in TOOLS

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

# Edited to KVB-only above, so a browser is sent to the KVB tool instead of Forbidden.
sent = clients["dpm"].get("/tool/dw")
assert sent.status_code == 302 and sent.headers["Location"].endswith("/kvb/tool/dw")
# Restore dpm as DPM-only through the admin endpoint.
c.post("/admin/users/dpm/access", headers=ADMIN, data={"can_dpm": "1"})
assert clients["dpm"].get("/tool/dw").status_code == 200
sent = clients["dpm"].get("/kvb/tool/dw")
assert sent.status_code == 302 and sent.headers["Location"].endswith("/tool/dw")
kvb_dw = clients["kvb"].get("/kvb/tool/dw")
assert kvb_dw.status_code == 200
assert b"KVB Plus workbook" in kvb_dw.data and b"translates KVB" in kvb_dw.data
assert b"Generate D&amp;W \xc2\xb7 KVB Tools" in kvb_dw.data
sent = clients["kvb"].get("/tool/dw")
assert sent.status_code == 302 and sent.headers["Location"].endswith("/kvb/tool/dw")
sent = clients["kvb"].get("/tool/segregate")
assert sent.status_code == 302 and sent.headers["Location"].endswith("/kvb/tool/segregate")
sent = clients["kvb"].get("/tool/pl-desk")
assert sent.status_code == 302 and sent.headers["Location"].rstrip("/").endswith("/kvb")
assert clients["kvb"].get("/kvb/tool/pl-desk").status_code == 404
sent = clients["dpm"].get("/kvb/tool/segregate")
assert sent.status_code == 302 and sent.headers["Location"].endswith("/tool/segregate")
kvb_seg = clients["kvb"].get("/kvb/tool/segregate")
assert kvb_seg.status_code == 200
assert b"Deal Segregator \xc2\xb7 KVB Tools" in kvb_seg.data
assert b'id="deals-drop"' in kvb_seg.data and b'id="equity-drop"' in kvb_seg.data
assert b"separate sheets" in kvb_seg.data
assert b'name="bulan"' not in kvb_seg.data
assert b"KVB Plus workbook" not in kvb_seg.data
assert b"translates KVB" not in kvb_seg.data
dpm_seg = clients["dpm"].get("/tool/segregate")
assert dpm_seg.status_code == 200
assert b"Deal Segregator \xc2\xb7 Dupoin DPM Tools" in dpm_seg.data
assert clients["both"].get("/tool/segregate").status_code == 200
assert clients["both"].get("/kvb/tool/segregate").status_code == 200
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
assert b"Deal Segregator" in kvb_page and b'href="/kvb/tool/segregate"' in kvb_page
assert b"PL by Desk" not in kvb_page
assert b'data-company="dpm" aria-disabled="true"' in kvb_page

# A KVB user can run Deal Segregator. The result stays inside KVB.
deals = (
    "Deal,Login,Time,Type,Entry,Symbol,Volume,Commission,Fee,Swap,Profit,Currency\n"
    "kvb-1,100,2026.07.01 00:00:00,buy,out,EURUSD,1,-1,0,0,10,USD\n"
).encode()
started = clients["kvb"].post(
    "/kvb/tool/segregate",
    data={"file": (io.BytesIO(deals), "KVB Deals.csv")},
    content_type="multipart/form-data")
assert started.status_code == 303, (started.status_code, started.data[:300])
assert started.headers["Location"].startswith("/kvb/job/")
kvb_job = started.headers["Location"].rstrip("/").rsplit("/", 1)[-1]
kvb_state = None
for _ in range(200):
    kvb_state = _read_state(kvb_job)
    if kvb_state and kvb_state["state"] in ("done", "failed"):
        break
    time.sleep(0.05)
assert kvb_state and kvb_state["state"] == "done", kvb_state
assert kvb_state["company"] == "kvb" and kvb_state["slug"] == "segregate"
assert kvb_state["name"] == "KVB Deals-hasil.zip"
kvb_job_page = clients["kvb"].get(f"/kvb/job/{kvb_job}")
assert kvb_job_page.status_code == 200
assert b"Processing \xc2\xb7 KVB Tools" in kvb_job_page.data
assert clients["kvb"].get(f"/job/{kvb_job}").status_code == 200
for path in (f"/kvb/job/{kvb_job}", f"/job/{kvb_job}"):
    sent = clients["dpm"].get(path)
    landed = sent.headers.get("Location", "")
    assert sent.status_code == 302 and landed.endswith("/") and "/kvb" not in landed, (
        path, sent.status_code, landed)
    assert kvb_job not in landed and b"KVB Deals-hasil.zip" not in sent.data
assert clients["dpm"].get(f"/kvb/job/{kvb_job}/download").status_code == 403
blocked = clients["dpm"].get(f"/job/{kvb_job}/download")
assert blocked.status_code == 403 and b"PK" not in blocked.data
ready = clients["kvb"].get(f"/kvb/job/{kvb_job}/download")
assert ready.status_code == 200
assert "KVB Deals-hasil.zip" in ready.headers["Content-Disposition"]
with zipfile.ZipFile(io.BytesIO(ready.data)) as archive:
    assert archive.namelist() == ["Deals - Daily.xlsx", "Deals - Monthly Summary.xlsx"]

# The other direction: a KVB user cannot open or download a DPM job.
dpm_job = "dpm-only-segregate-job"
secret = b"DPM-SECRET-RESULT"
(JOBS / dpm_job).mkdir(exist_ok=True)
(JOBS / dpm_job / "hasil.zip").write_bytes(secret)
_write_state(dpm_job, state="done", name="dpm-hasil.zip", started=time.time(),
             error=None, log=None, path=str(JOBS / dpm_job / "hasil.zip"),
             slug="segregate", company="dpm")
assert clients["dpm"].get(f"/job/{dpm_job}").status_code == 200
for path in (f"/job/{dpm_job}", f"/kvb/job/{dpm_job}"):
    sent = clients["kvb"].get(path)
    assert sent.status_code == 302 and sent.headers["Location"].rstrip("/").endswith("/kvb"), path
    assert secret not in sent.data and dpm_job not in sent.headers["Location"]
stolen = clients["kvb"].get(f"/job/{dpm_job}/download")
assert stolen.status_code == 403 and secret not in stolen.data
stolen_kvb_url = clients["kvb"].get(f"/kvb/job/{dpm_job}/download")
assert stolen_kvb_url.status_code == 403 and secret not in stolen_kvb_url.data

admin_page = c.get("/admin", headers=ADMIN).data
assert b'name="can_dpm"' in admin_page and b'name="can_kvb"' in admin_page
assert b"/admin/users/legacy-user/access" in admin_page

print("OK company access")
