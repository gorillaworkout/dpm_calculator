#!/usr/bin/env python3
"""Visible version: VERSION file, footer on every page, error/STOP text, /version.

python3 test_version.py
"""
import base64
import json
import os
import re
import tempfile
import uuid
from pathlib import Path

TMP = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TMP / "t.db")
os.environ["DW_JOB_DIR"] = str(TMP / "jobs")
os.environ["DW_ADMIN_USER"] = "boss"
os.environ["DW_ADMIN_PASS"] = "bosspassword1"

import app as webapp  # noqa: E402
from app import app  # noqa: E402

app.config["TESTING"] = True
ROOT = Path(__file__).resolve().parent
HTML = {"Accept": "text/html,application/xhtml+xml"}


def basic(user, pw):
    return {"Authorization": "Basic " + base64.b64encode(f"{user}:{pw}".encode()).decode(),
            **HTML}


BOSS = basic("boss", "bosspassword1")

# --- VERSION file -------------------------------------------------------------
v = webapp.baca_versi(ROOT / "VERSION")
assert re.fullmatch(r"1\.\d+", v["version"]), v
assert int(v["version"].split(".")[1]) >= 22, v
assert re.fullmatch(r"\d{4}-\d{2}-\d{2}", v["build"]), v
assert webapp.VERSI == v, "app must read VERSION at startup"
LABEL = webapp.VERSI_LABEL
assert LABEL == f"DW Calculator v{v['version']} · {v['build']}", LABEL
assert "Bump" in (ROOT / "VERSION").read_text(), "VERSION must say how to bump it"
assert not re.search(r"\b[0-9a-f]{40}\b", (ROOT / "VERSION").read_text()), "no SHA in VERSION"

p = TMP / "V1"
assert webapp.baca_versi(TMP / "missing") == {"version": "unknown", "build": ""}
assert webapp.label_versi({"version": "unknown", "build": ""}) == "DW Calculator version unknown"
p.write_text("1.31\n", encoding="utf-8")
assert webapp.baca_versi(p) == {"version": "1.31", "build": ""}
assert webapp.label_versi(webapp.baca_versi(p)) == "DW Calculator v1.31"
p.write_text("# c\nversion = 1.32\nbuild= 2026-11-01\nother=x\n", encoding="utf-8")
assert webapp.baca_versi(p) == {"version": "1.32", "build": "2026-11-01"}
p.write_bytes(b"\xff\xfe\x00bad")
assert webapp.baca_versi(p)["version"] == "unknown"
print("OK VERSION file")

LB = LABEL.encode()
LB_HTML = LABEL.replace("·", "&middot;").encode()


def has_label(data):
    return LB in data or LB_HTML in data


# --- /version: login required --------------------------------------------------
c = app.test_client()
r = c.get("/version")
assert r.status_code == 401, r.status_code
r = c.get("/version", headers=HTML)
assert r.status_code == 302 and "/login" in r.headers["Location"]
r = c.get("/version", headers={"Authorization": BOSS["Authorization"]})
assert r.status_code == 200 and r.is_json, r.data[:200]
assert r.get_json() == {"name": "DW Calculator", "version": v["version"],
                        "build": v["build"], "label": LABEL}
print("OK /version")

# --- footer on every page, sign-in page included -------------------------------
r = app.test_client().get("/login")
assert r.status_code == 200 and has_label(r.data) and b'id="app-version"' in r.data
r = app.test_client().post("/login", data={"username": "boss", "password": "no"})
assert r.status_code == 401 and b"Wrong username or password" in r.data and has_label(r.data)
for path in ("/", "/kvb", "/tool/dw", "/tool/segregate", "/kvb/tool/segregate", "/admin"):
    r = c.get(path, headers=BOSS)
    assert r.status_code == 200, (path, r.status_code)
    assert b'id="app-version"' in r.data and has_label(r.data), path
print("OK footer on every page")

# --- flash message --------------------------------------------------------------
r = c.post("/admin/users", headers=BOSS,
           data={"username": "dpmonly", "password": "dpmonly-password-1", "can_dpm": "1"})
assert r.status_code == 302, r.data[:300]
d = app.test_client()
r = d.get("/kvb/tool/segregate", headers=basic("dpmonly", "dpmonly-password-1"))
assert r.status_code == 302
r = d.get(r.headers["Location"], headers=basic("dpmonly", "dpmonly-password-1"))
assert b"does not have access to KVB" in r.data
note = re.search(rb'<div class="note warn"[^>]*>(.*?)</div>', r.data, re.S).group(1)
assert has_label(note), note
print("OK flash message")

# --- failed job page: STOP note and the $ deal_segregator.py log box -------------
jid = uuid.uuid4().hex
webapp._write_state(jid, state="failed", slug="segregate", company="kvb", name="x.zip",
                    started=0, version="DW Calculator v1.21 · 2026-10-18",
                    error="STOP: required columns are missing in 'a.csv'",
                    log="$ deal_segregator.py\nSTOP: required columns are missing in 'a.csv'")
r = c.get(f"/kvb/job/{jid}", headers=BOSS)
assert r.status_code == 422
pre = re.search(rb"<pre>(.*?)</pre>", r.data, re.S).group(1)
assert has_label(pre.split(b"\n")[0]), pre[:200]
assert b"job ran on DW Calculator v1.21" in pre and b"$ deal_segregator.py" in pre
notes = re.findall(rb'<div class="note bad">(.*?)</div>', r.data, re.S)
assert notes and all(has_label(n) for n in notes), notes
# same version: no "job ran on" remark
webapp._write_state(jid, version=LABEL)
r = c.get(f"/kvb/job/{jid}", headers=BOSS)
assert b"job ran on" not in r.data
# a new job records the version it runs on
src = (ROOT / "app.py").read_text(encoding="utf-8")
assert "version=VERSI_LABEL" in src
print("OK job error page")

# --- 410 (collected) page --------------------------------------------------------
jid2 = uuid.uuid4().hex
webapp._write_state(jid2, state="collected", slug="segregate", company="dpm", name="r.zip",
                    started=0)
r = c.get(f"/job/{jid2}/download", headers=BOSS)
assert r.status_code == 410 and has_label(r.data), r.status_code
print("OK 410 page")

# --- chunked-upload error text in tool.html ------------------------------------
r = c.get("/kvb/tool/segregate", headers=BOSS)
assert ("var APP_VERSION = " + json.dumps(LABEL)).encode() in r.data
assert b"state === 'failed' ? detail + ' [' + APP_VERSION + ']'" in r.data
print("OK upload error text")

# --- 404 / 403 / 405 / 500 pages ---------------------------------------------------
r = c.get("/tool/nope", headers=BOSS)
assert r.status_code == 404 and b"Not Found" in r.data and has_label(r.data)
r = d.get("/kvb/tool/segregate", headers={"Authorization": basic("dpmonly", "dpmonly-password-1")["Authorization"]})
assert r.status_code == 403 and b"does not have access to KVB" in r.data and has_label(r.data)
r = c.put("/version", headers=BOSS)
assert r.status_code == 405 and "Allow" in r.headers and has_label(r.data)
from werkzeug.exceptions import InternalServerError  # noqa: E402
with app.test_request_context("/"):
    resp = webapp._halaman_error(InternalServerError())
    assert resp.status_code == 500 and has_label(resp.get_data())
print("OK error pages")
print("OK version")
