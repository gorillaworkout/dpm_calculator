#!/usr/bin/env python3
"""Browser visits to the wrong company go to a page the user can use.

python3 test_company_redirect.py
"""
import base64
import os
import sqlite3
import tempfile
import time
from pathlib import Path
from urllib.parse import urlparse

TMP = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TMP / "users.db")
os.environ["DW_JOB_DIR"] = str(TMP / "jobs")
os.environ["DW_ADMIN_USER"] = "redirect-admin"
os.environ["DW_ADMIN_PASS"] = "redirect-admin-password-1"

from app import JOBS, _write_state, app  # noqa: E402

app.config["TESTING"] = True
HTML = {"Accept": "text/html,application/xhtml+xml"}
STAR = {"Accept": "*/*"}


def sign_in(username, password):
    client = app.test_client()
    response = client.post("/login", data={"username": username, "password": password})
    assert response.status_code == 302, (username, response.status_code, response.data[:200])
    return client


def location(response):
    return response.headers.get("Location", "")


admin = sign_in("redirect-admin", "redirect-admin-password-1")
for name, access in (
    ("dpm-only", {"can_dpm": "1"}),
    ("kvb-only", {"can_kvb": "1"}),
    ("both", {"can_dpm": "1", "can_kvb": "1"}),
    ("no-access", {"can_dpm": "1"}),
):
    created = admin.post("/admin/users", data={
        "username": name, "password": f"{name}-password-1", **access})
    assert created.status_code == 302, (name, created.status_code, created.data[:200])

with sqlite3.connect(os.environ["DW_AUTH_DB"]) as con:
    con.execute("UPDATE users SET can_dpm=0, can_kvb=0 WHERE username='no-access'")

dpm = sign_in("dpm-only", "dpm-only-password-1")
kvb = sign_in("kvb-only", "kvb-only-password-1")
both = sign_in("both", "both-password-1")
none = sign_in("no-access", "no-access-password-1")
anon = app.test_client()

# Unauthenticated browser GET still goes to the login page.
for path in ("/", "/tool/segregate", "/tool/pl-desk", "/kvb", "/kvb/tool/segregate"):
    response = anon.get(path, headers=HTML)
    assert response.status_code == 302, (path, response.status_code)
    assert "/login" in location(response), (path, location(response))
    assert f"next={path}" in location(response) or f"next={path.replace('/', '%2F')}" in location(response), location(response)
    assert "WWW-Authenticate" not in response.headers

# Accept */* and a missing Accept stay a plain 401. curl and fetch send */*;
# the chunked uploader only POSTs, so it is covered by the POST check below.
for headers in (None, STAR):
    response = anon.get("/tool/segregate", headers=headers)
    assert response.status_code == 401, (headers, response.status_code, location(response))
    assert "WWW-Authenticate" in response.headers
    assert b"Forbidden" not in response.data
response = anon.post("/upload/begin", headers=HTML)
assert response.status_code == 401
response = anon.post("/tool/segregate", headers=HTML)
assert response.status_code == 401

# A KVB-only browser is sent to the matching KVB page, with a notice.
sent = kvb.get("/tool/segregate", headers=HTML)
assert sent.status_code == 302, (sent.status_code, sent.data[:200])
assert location(sent).endswith("/kvb/tool/segregate"), location(sent)
page = kvb.get("/tool/segregate", headers=HTML, follow_redirects=True)
assert page.status_code == 200
assert b"Opened the KVB page instead." in page.data
assert b"does not have access to DPM" in page.data
assert b"Deal Segregator" in page.data
assert b"Forbidden" not in page.data

sent = kvb.get("/", headers=HTML)
assert sent.status_code == 302 and location(sent).rstrip("/").endswith("/kvb"), location(sent)
home = kvb.get("/", headers=HTML, follow_redirects=True)
assert home.status_code == 200 and b"KVB Tools" in home.data
assert b"Forbidden" not in home.data

# PL by Desk exists only on DPM, so a KVB user lands on the KVB home.
sent = kvb.get("/tool/pl-desk", headers=HTML)
assert sent.status_code == 302 and location(sent).rstrip("/").endswith("/kvb"), location(sent)
assert "/tool/pl-desk" not in location(sent)
desk = kvb.get("/tool/pl-desk", headers=HTML, follow_redirects=True)
assert desk.status_code == 200 and b"KVB Tools" in desk.data
assert b"FinanceOS" not in desk.data and b"Forbidden" not in desk.data

# The other direction: a DPM-only user reaches the DPM segregator.
sent = dpm.get("/kvb/tool/segregate", headers=HTML)
assert sent.status_code == 302 and location(sent).endswith("/tool/segregate"), location(sent)
assert "/kvb/" not in location(sent)
page = dpm.get("/kvb/tool/segregate", headers=HTML, follow_redirects=True)
assert page.status_code == 200
assert b"Opened the DPM page instead." in page.data
assert b"Dupoin DPM Tools" in page.data
assert b"Forbidden" not in page.data

sent = dpm.get("/kvb", headers=HTML)
assert sent.status_code == 302, (sent.status_code, location(sent))
# The DPM home is "/", not a KVB URL.
assert urlparse(location(sent)).path == "/"

# Users who can open the page are left there.
for client, path in (
    (dpm, "/"), (dpm, "/tool/segregate"), (dpm, "/tool/pl-desk"),
    (kvb, "/kvb"), (kvb, "/kvb/tool/segregate"),
    (both, "/"), (both, "/tool/segregate"), (both, "/tool/pl-desk"),
    (both, "/kvb"), (both, "/kvb/tool/segregate"),
):
    response = client.get(path, headers=HTML)
    assert response.status_code == 200, (path, response.status_code, location(response))
    assert b"Forbidden" not in response.data

# No company access: the login page explains it and offers a switch.
sent = none.get("/tool/segregate", headers=HTML)
assert sent.status_code == 302 and "/login" in location(sent), location(sent)
login_page = none.get("/tool/segregate", headers=HTML, follow_redirects=True)
assert login_page.status_code == 200
assert b"Your account does not have access to DPM. Please sign in with a DPM account." in login_page.data
assert b"Sign in with a different account" in login_page.data
assert b"/logout" in login_page.data and b"next=" in login_page.data
assert b"Forbidden" not in login_page.data
sent = none.get("/kvb", headers=HTML, follow_redirects=True)
assert b"Your account does not have access to KVB. Please sign in with a KVB account." in sent.data

# Switching account signs out and keeps the original page for the next sign-in.
switch = none.get("/logout", query_string={"next": "/tool/segregate"})
assert switch.status_code == 302 and "/login" in location(switch)
assert "next=/tool/segregate" in location(switch) or "next=%2Ftool%2Fsegregate" in location(switch)
assert none.get("/", headers=HTML).status_code == 302, "logout did not clear the session"

# Form posts and the uploader stay on 401/403. A redirect would drop the body.
posted = kvb.post("/tool/segregate", headers=HTML, data={})
assert posted.status_code == 403, posted.status_code
assert posted.status_code != 302
assert kvb.post("/upload/begin", headers=HTML, data={"bytes": "1"}).status_code != 302

# A cross-company job page redirects home and does not reveal the result.
secret = b"DPM-SECRET-RESULT-BYTES"
job_id = "dpm-secret-job"
(JOBS / job_id).mkdir(parents=True, exist_ok=True)
(JOBS / job_id / "hasil.zip").write_bytes(secret)
_write_state(job_id, state="done", name="dpm-secret-hasil.zip", started=time.time(),
             error=None, log="private log line", path=str(JOBS / job_id / "hasil.zip"),
             slug="segregate", company="dpm")
for path in (f"/job/{job_id}", f"/kvb/job/{job_id}"):
    response = kvb.get(path, headers=HTML)
    assert response.status_code == 302, (path, response.status_code, response.data[:200])
    assert location(response).rstrip("/").endswith("/kvb"), (path, location(response))
    assert job_id not in location(response)
    assert secret not in response.data
    assert b"dpm-secret-hasil.zip" not in response.data
    assert b"private log line" not in response.data
    followed = kvb.get(path, headers=HTML, follow_redirects=True)
    assert followed.status_code == 200
    assert secret not in followed.data
    assert b"dpm-secret-hasil.zip" not in followed.data
    assert b"private log line" not in followed.data
    download = kvb.get(path + "/download", headers=HTML)
    assert download.status_code == 403, (path, download.status_code, location(download))
    assert secret not in download.data
    assert b"PK" not in download.data

# The owner still sees the job. The other company's download stays closed.
assert dpm.get(f"/job/{job_id}", headers=HTML).status_code == 200
owned = dpm.get(f"/job/{job_id}/download", headers=HTML)
assert owned.status_code == 200 and secret in owned.data

# A script Accept does not turn the job page into a redirect that leaks HTML.
script = kvb.get(f"/job/{job_id}", headers=STAR)
assert script.status_code == 403, (script.status_code, location(script))
assert secret not in script.data

# Admin stays a 403 for a signed-in user who is not an admin.
assert kvb.get("/admin", headers=HTML).status_code == 403
assert anon.get("/admin", headers=HTML).status_code == 302

# Basic Auth plus a browser Accept follows the same redirect.
token = base64.b64encode(b"kvb-only:kvb-only-password-1").decode()
basic = app.test_client().get(
    "/tool/segregate",
    headers={"Authorization": f"Basic {token}", "Accept": "text/html"})
assert basic.status_code == 302 and location(basic).endswith("/kvb/tool/segregate"), location(basic)

print("OK company redirect")
