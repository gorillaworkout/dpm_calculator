#!/usr/bin/env python3
"""Upload queue ownership (PR #25: reports from Bayu and sulin on v1.24).

Root cause: /upload/begin treated every live staging upload as someone else's.
A user's own stalled or abandoned upload of 1 GiB or more kept a fresh heartbeat
for 15 minutes, so their next upload was queued behind themselves.

python3 test_upload_queue_owner.py
"""
import base64
import io
import json
import os
import tempfile
import time
import uuid
from collections import namedtuple
from pathlib import Path

TEST_ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TEST_ROOT / "t.db")
os.environ["DW_JOB_DIR"] = str(TEST_ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "admin"
os.environ["DW_ADMIN_PASS"] = "admin-password-1"

import chunked  # noqa: E402
import cleanup  # noqa: E402
import app as web  # noqa: E402
from app import JOBS, app  # noqa: E402

app.config["TESTING"] = True
GIB = 1024 ** 3
Usage = namedtuple("Usage", "total used free")
chunked.shutil.disk_usage = lambda p: Usage(100 * GIB, 0, 60 * GIB)   # plenty of disk


def client(user, pw):
    c = app.test_client()
    c.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(f"{user}:{pw}".encode()).decode()
    return c


admin = client("admin", "admin-password-1")
for name in ("bayu", "alice"):
    r = admin.post("/admin/users", data={"username": name, "password": f"{name}-secret-pw-1",
                                         "can_dpm": "1", "can_kvb": "1"})
    assert r.status_code == 302, r.status_code
bayu = client("bayu", "bayu-secret-pw-1")
alice = client("alice", "alice-secret-pw-1")
SECRETS = ("bayu-secret-pw-1", "alice-secret-pw-1", "admin-password-1")


def begin(c, nbytes, ticket=None):
    data = {"bytes": str(nbytes)}
    if ticket:
        data["ticket"] = ticket
    return c.post("/upload/begin", data=data)


def manifest(uid):
    return json.loads((JOBS / "staging" / uid / "manifest.json").read_text())


def set_heartbeat(uid, ago):
    m = manifest(uid)
    m["last_chunk"] = time.time() - ago
    (JOBS / "staging" / uid / "manifest.json").write_text(json.dumps(m))


def chunk(c, uid, blob=b"abc", seq=0):
    return c.post("/upload/chunk", data={"id": uid, "name": "a.csv", "index": "0", "seq": str(seq),
                                         "size": str(len(blob)), "chunk": (io.BytesIO(blob), "blob")},
                  content_type="multipart/form-data")


def no_secrets(r):
    body = r.get_data(as_text=True)
    assert not any(s in body for s in SECRETS), body


def reset():
    for p in (JOBS / "staging").iterdir():
        if p.is_dir():
            chunked.shutil.rmtree(p)
        else:
            p.unlink()
    for p in (JOBS / "queue").iterdir():
        p.unlink()


# --- 1. own stalled heavy upload: superseded, not queued -------------------------
r = begin(bayu, 2 * GIB)
assert r.status_code == 200, r.get_data(as_text=True)
first = r.get_json()["id"]
assert manifest(first)["user"] == "bayu", "owner stored at begin"
assert chunk(bayu, first).status_code == 200
r = begin(bayu, 2 * GIB)                       # Bayu uploads again
assert r.status_code == 200, ("own upload queued me", r.status_code, r.get_json())
second = r.get_json()["id"]
assert not (JOBS / "staging" / first).exists(), "own older upload not discarded"
r = chunk(bayu, first, seq=1)                   # the older tab keeps sending
assert r.status_code == 410 and r.get_json()["superseded"], (r.status_code, r.get_json())
assert "replaced by a newer upload" in r.get_json()["error"]
print("OK own stalled upload is superseded")

# --- 2. another user's live heavy upload: queued, notice names the account ----------
r = begin(alice, 3 * GIB)
assert r.status_code == 202, r.status_code
q = r.get_json()
assert q["message"].startswith("Server sedang dipakai oleh bayu."), q["message"]
assert "uploading 2.0 GB" in q["message"]
assert q["blocked_by"] == [{"kind": "upload", "user": "bayu", "own": False}], q
no_secrets(r)
assert (JOBS / "staging" / second).is_dir(), "another user's upload was touched"
# Alice's poll does not take Bayu's data away either
assert begin(alice, 3 * GIB, q["ticket"]).status_code == 202
assert (JOBS / "staging" / second).is_dir()
print("OK other user's upload: queued with the owner named")

# --- 3. own queue slot superseded by a new attempt (another tab) ---------------------
r1 = begin(alice, 3 * GIB)                       # alice, second tab, fresh begin
assert r1.status_code == 202 and r1.get_json()["position"] == 1, r1.get_json()
r = begin(alice, 3 * GIB, q["ticket"])           # the first tab polls again
assert r.status_code == 409 and r.get_json()["superseded"], r.get_json()
print("OK own older queue ticket superseded")

# --- 4. stale owner: no chunk for > 5 min stops blocking; data kept until 15 min ------
set_heartbeat(second, cleanup.UPLOAD_BLOCK_SECONDS + 30)
r = begin(alice, 3 * GIB, r1.get_json()["ticket"])
assert r.status_code == 200, r.get_json()
alice_up = r.get_json()["id"]
assert (JOBS / "staging" / second).is_dir(), "idle (not expired) upload deleted too early"
set_heartbeat(second, cleanup.UPLOAD_IDLE_SECONDS + 30)
assert begin(admin, 10).status_code == 200
assert not (JOBS / "staging" / second).exists(), "expired upload not cleaned"
print("OK stale upload stops blocking at 5 min, is removed at 15 min")

# --- 5. discard only your own upload -------------------------------------------------
r = bayu.post("/upload/discard", data={"id": alice_up})
assert r.status_code == 403 and (JOBS / "staging" / alice_up).is_dir()
r = alice.post("/upload/discard", data={"id": alice_up, "reason": "pagehide"})
assert r.status_code == 200 and not (JOBS / "staging" / alice_up).exists()
reset()
print("OK discard is owner-only")

# --- 6. legacy manifest without owner: generic notice -------------------------------
legacy = begin(admin, 2 * GIB).get_json()["id"]
m = manifest(legacy)
m.pop("user")
(JOBS / "staging" / legacy / "manifest.json").write_text(json.dumps(m))
r = begin(bayu, 2 * GIB)
assert r.status_code == 202
msg = r.get_json()["message"]
assert "dipakai oleh" not in msg and "processing another user's file" in msg, msg
assert (JOBS / "staging" / legacy).is_dir(), "legacy upload of unknown owner removed"
assert bayu.post("/upload/discard", data={"id": legacy}).status_code == 200   # legacy: allowed
reset()
print("OK legacy entry: generic notice")


# --- 7. own RUNNING job: not cancelled; clear message + link ---------------------------
def running_job(user, company="kvb"):
    jid = uuid.uuid4().hex
    (JOBS / jid).mkdir()
    web._write_state(jid, state="running", name="Deals-hasil.zip", slug="segregate",
                     company=company, user=user, started=time.time() - 120)
    return jid, cleanup.acquire_job_lock(JOBS, jid)


jid, lock = running_job("bayu")
r = begin(bayu, 2 * GIB)
assert r.status_code == 202, r.status_code
j = r.get_json()
assert j["message"].startswith("Your previous run (Deals-hasil.zip) is still processing"), j["message"]
assert "cannot be stopped safely" in j["message"] and j["job_url"] == f"/kvb/job/{jid}"
assert j["blocked_by"] == [{"kind": "job", "user": "bayu", "own": True}]
assert cleanup.job_locked(JOBS, jid) and (JOBS / jid).is_dir(), "own running job was touched"
# alice sees bayu named, and no link to his job
r = begin(alice, 2 * GIB)
assert r.status_code == 202 and r.get_json()["job_url"] is None
assert r.get_json()["message"].startswith("Server sedang dipakai oleh bayu."), r.get_json()["message"]
assert "Deal Segregator run, processing since" in r.get_json()["message"]
assert "Deals-hasil.zip" not in r.get_json()["message"], "file name of another user leaked"
no_secrets(r)
cleanup.release_job_lock(lock)
reset()
# small uploads are not held back by a running job when disk is fine (unchanged rule)
jid, lock = running_job("alice")
assert begin(bayu, 1000).status_code == 200
cleanup.release_job_lock(lock)
reset()
print("OK own running job: message and link, never cancelled")

# --- 8. staged upload of another account cannot be processed by you --------------------
up = begin(alice, 3).get_json()["id"]
assert chunk(alice, up).status_code == 200
r = bayu.post("/kvb/tool/segregate", data={"upload_id": up}, content_type="multipart/form-data",
              headers={"Accept": "text/html"})
assert r.status_code == 400 and b"belongs to another account" in r.data
assert (JOBS / "staging" / up).is_dir()
reset()
print("OK upload_id of another account refused")

# --- 9. the new job state records its owner -------------------------------------------
src = Path(web.__file__).read_text(encoding="utf-8")
assert 'user=request.user["username"]' in src

# --- 10. limits unchanged; queue tickets keep TTL; legacy tickets still read -----------
assert chunked.MAX_TOTAL == 12 * GIB and chunked.MAX_FILES == 70
assert begin(bayu, 12 * GIB + 1).status_code == 413
assert cleanup.QUEUE_TICKET_TTL == 30 and cleanup.UPLOAD_IDLE_SECONDS == 15 * 60
t = "c" * 32
(JOBS / "queue" / t).write_text(str(time.time()))          # pre-#25 ticket format
assert cleanup.queue_entries(JOBS)[0][1:] == (t, None)
old = time.time() - 120
os.utime(JOBS / "queue" / t, (old, old))
assert cleanup.queue_entries(JOBS) == [] and not (JOBS / "queue" / t).exists()
reset()
print("OK limits and ticket expiry")

# --- 11. the request log says what happened -------------------------------------------
import logging  # noqa: E402
lines = []
h = logging.Handler()
h.emit = lambda rec: lines.append(rec.getMessage())
chunked.log.addHandler(h)
a = begin(bayu, 2 * GIB).get_json()["id"]
begin(bayu, 2 * GIB)
begin(alice, 2 * GIB)
assert any("user=bayu" in ln and "note=superseded own uploads 1" in ln for ln in lines), lines
assert any("user=alice" in ln and " 202 " in ln and "note=queued behind bayu:upload" in ln
           for ln in lines), lines
assert not (JOBS / "staging" / a).exists()
print("OK queue owner")
