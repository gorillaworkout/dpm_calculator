#!/usr/bin/env python3
"""Chunked upload checks. Run: python3 test_chunked.py

Proves the pieces are reassembled byte-for-byte and that the staging area cannot
be used to write outside itself.
"""
import base64
import hashlib
import io
import os
import tempfile
import time
from pathlib import Path

TEST_ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TEST_ROOT / "t.db")
os.environ["DW_JOB_DIR"] = str(TEST_ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "tester"
os.environ["DW_ADMIN_PASS"] = "tester-password-1"

import chunked  # noqa: E402
import app as web  # noqa: E402
from app import JOBS, app  # noqa: E402

assert chunked.MAX_TOTAL == 5 * 1024 * 1024 * 1024

app.config["TESTING"] = True
c = app.test_client()
c.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(
    b"tester:tester-password-1").decode()

CHUNK = 64 * 1024


def begin():
    r = c.post("/upload/begin")
    assert r.status_code == 200, r.status_code
    return r.get_json()["id"]


def send(uid, name, index, seq, blob):
    return c.post("/upload/chunk", data={
        "id": uid, "name": name, "index": str(index), "seq": str(seq),
        "chunk": (io.BytesIO(blob), "part"),
    }, content_type="multipart/form-data")


# --- the endpoints are behind auth too --------------------------------------
anon = app.test_client()
assert anon.post("/upload/begin").status_code == 401, "/upload/begin was open"
assert anon.post("/upload/chunk", data={}).status_code == 401, "/upload/chunk was open"

# --- start failures are explicit and leave no empty staging directory --------
original_write_manifest = chunked._write_manifest
try:
    def fail_manifest(*_args, **_kwargs):
        raise OSError(122, "Disk quota exceeded")
    chunked._write_manifest = fail_manifest
    (JOBS / "staging").mkdir(parents=True, exist_ok=True)
    before = {p.name for p in (JOBS / "staging").iterdir()}
    r = c.post("/upload/begin", data={"bytes": str(100 * 1024 * 1024)})
    assert r.status_code == 507, (r.status_code, r.data)
    assert b"temporary storage" in r.data and b"try again" in r.data
    assert {p.name for p in (JOBS / "staging").iterdir()} == before
finally:
    chunked._write_manifest = original_write_manifest

# --- browser can discard an interrupted upload immediately ------------------
uid = begin()
stage_dir = JOBS / "staging" / uid
assert c.post("/upload/discard", data={"id": uid}).status_code == 200
assert not stage_dir.exists(), "discard endpoint left interrupted upload data"

# --- a file split into pieces is reassembled exactly ------------------------
payload = os.urandom(CHUNK * 3 + 517)
digest = hashlib.sha256(payload).hexdigest()
uid = begin()
seq = 0
for off in range(0, len(payload), CHUNK):
    r = send(uid, "deals.csv", 0, seq, payload[off:off + CHUNK])
    assert r.status_code == 200, (r.status_code, r.get_json())
    seq += 1

staged = chunked.staged_files(JOBS, uid)
assert staged and len(staged) == 1, staged
path, original = staged[0]
assert original == "deals.csv"
assert hashlib.sha256(path.read_bytes()).hexdigest() == digest, "reassembled bytes differ"

# --- 20 files keep their own identity and order -----------------------------
uid = begin()
payloads = [f"row-{i}".encode() for i in range(20)]
for i, payload_i in enumerate(payloads):
    assert send(uid, f"day-{i + 1}.csv", i, 0, payload_i).status_code == 200
staged = chunked.staged_files(JOBS, uid)
assert [n for _p, n in staged] == [f"day-{i}.csv" for i in range(1, 21)], staged
assert [p.read_bytes() for p, _n in staged] == payloads

# --- retrying an accepted chunk is idempotent -------------------------------
retry_uid = begin()
assert send(retry_uid, "x.csv", 0, 0, b"aaa").status_code == 200
r = send(retry_uid, "x.csv", 0, 0, b"aaa")
assert r.status_code == 200, (r.status_code, r.get_json())
assert chunked.staged_files(JOBS, retry_uid)[0][0].read_bytes() == b"aaa"

# --- retry at the exact total ceiling preserves the completed upload --------
old_max = chunked.MAX_TOTAL
chunked.MAX_TOTAL = 6
try:
    uid = begin()
    assert send(uid, "limit.csv", 0, 0, b"aaa").status_code == 200
    assert send(uid, "limit.csv", 0, 1, b"bbb").status_code == 200
    r = send(uid, "limit.csv", 0, 1, b"bbb")
    assert r.status_code == 200, (r.status_code, r.get_json())
    assert chunked.staged_files(JOBS, uid)[0][0].read_bytes() == b"aaabbb"
finally:
    chunked.MAX_TOTAL = old_max

# --- genuinely out-of-order chunks are refused, not silently corrupted ------
r = send(retry_uid, "x.csv", 0, 5, b"bbb")
assert r.status_code == 409, r.status_code
assert chunked.staged_files(JOBS, retry_uid)[0][0].read_bytes() == b"aaa"

# --- rejected file types ----------------------------------------------------
uid = begin()
r = send(uid, "payload.exe", 0, 0, b"MZ")
assert r.status_code == 400 and b"not a .csv" in r.data

# --- path traversal cannot escape the staging directory ---------------------
uid = begin()
r = send(uid, "../../../../etc/passwd.csv", 0, 0, b"root:x:0:0")
assert r.status_code == 200, "the name should be sanitised, not rejected outright"
staged = chunked.staged_files(JOBS, uid)
stored = staged[0][0].resolve()
assert (JOBS / "staging").resolve() in stored.parents, f"escaped staging: {stored}"

# a forged upload id must not resolve to anything
assert chunked.staged_files(JOBS, "../../etc") is None
assert chunked.staged_files(JOBS, "zz" * 16) is None
assert c.post("/upload/chunk", data={"id": "../../etc", "name": "a.csv",
                                     "index": "0", "seq": "0",
                                     "chunk": (io.BytesIO(b"x"), "p")},
              content_type="multipart/form-data").status_code == 404

# --- an unknown id is a clean validation error, never a 500 -----------------
r = c.post("/tool/segregate", data={"upload_id": "ab" * 16},
           content_type="multipart/form-data")
assert r.status_code == 400 and b"expired" in r.data, r.status_code

# --- a completed run consumes the staged files ------------------------------
uid = begin()
assert send(uid, "small.csv", 0, 0, b"Deal,Login\n1,2\n").status_code == 200
stage_dir = (JOBS / "staging" / uid)
assert stage_dir.is_dir()
chunked.discard(JOBS, uid)
assert not stage_dir.exists(), "staging survived discard"

# --- the sweeper removes abandoned uploads ----------------------------------
uid = begin()
stage_dir = JOBS / "staging" / uid
old = time.time() - 7 * 3600
os.utime(stage_dir, (old, old))
os.utime(stage_dir / "manifest.json", (old, old))
chunked.sweep(JOBS)
assert not stage_dir.exists(), "an abandoned upload was left on disk"

# Directory mtime does not change while a chunk is appended. A recently updated
# manifest is the activity marker and must protect a long-running upload.
uid = begin()
stage_dir = JOBS / "staging" / uid
os.utime(stage_dir, (old, old))
chunked.sweep(JOBS)
assert stage_dir.exists(), "an active long upload was deleted"
chunked.discard(JOBS, uid)

# The generic completed-job sweeper must never delete active chunk staging.
uid = begin()
stage_dir = JOBS / "staging" / uid
os.utime(JOBS / "staging", (old, old))
web._sweep_old_jobs()
assert stage_dir.exists(), "generic job sweep deleted active chunk staging"
chunked.discard(JOBS, uid)

print("OK chunked")
