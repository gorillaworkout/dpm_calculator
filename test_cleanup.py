#!/usr/bin/env python3
"""Cleanup-on-new-run and heavy-job queue checks. Run: python3 test_cleanup.py"""
import base64
import json
import os
import tempfile
import time
from collections import namedtuple
from pathlib import Path

TEST_ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TEST_ROOT / "t.db")
os.environ["DW_JOB_DIR"] = str(TEST_ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "tester"
os.environ["DW_ADMIN_PASS"] = "tester-password-1"

import chunked  # noqa: E402
import cleanup  # noqa: E402
import app as web  # noqa: E402
from app import JOBS, app  # noqa: E402

app.config["TESTING"] = True
c = app.test_client()
c.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(
    b"tester:tester-password-1").decode()
OLD = time.time() - 3 * 3600
_GIB = 1024 ** 3


def age(p):
    for root, dirs, files in os.walk(p, topdown=False):
        for n in dirs + files:
            os.utime(os.path.join(root, n), (OLD, OLD))
    os.utime(p, (OLD, OLD))


def fake_job(state, size=4096, slug="segregate"):
    import uuid
    jid = uuid.uuid4().hex
    d = JOBS / jid
    d.mkdir()
    (d / "input.csv").write_bytes(b"x" * size)
    (d / "deals.sqlite").write_bytes(b"i" * size)
    (d / "hasil.zip").write_bytes(b"r" * 10)
    web._write_state(jid, state=state, name="r.zip", slug=slug, company="dpm",
                     started=OLD, path=str(d / "hasil.zip"), error=None, log=None)
    age(d)
    age(JOBS / f"{jid}.json")
    return jid


def old_stage():
    uid = c.post("/upload/begin", data={"bytes": "10"}).get_json()["id"]
    st = JOBS / "staging" / uid
    m = json.loads((st / "manifest.json").read_text())
    m["last_chunk"] = OLD
    (st / "manifest.json").write_text(json.dumps(m))
    (st / "part-00.csv").write_bytes(b"s" * 10)
    age(st)
    return st


# 1. Finished/abandoned job dirs, idle staging and temp files go on upload init;
#    a locked (running) job is untouched.
done = fake_job("done")
failed = fake_job("failed")
running = fake_job("running")
lock = cleanup.acquire_job_lock(JOBS, running)
age(JOBS / f"{running}.lock")
temp = JOBS / "stray.tmp"
temp.write_bytes(b"t" * 100)
age(temp)
st_old = old_stage()
r = c.post("/upload/begin", data={"bytes": "10"})
assert r.status_code == 200, r.get_data(as_text=True)
assert r.get_json()["freed"] > 0
assert not (JOBS / done).exists() and not (JOBS / failed).exists()
assert not temp.exists() and not st_old.exists(), "old temp/staging survived"
assert (JOBS / running / "input.csv").is_file(), "an active job was deleted"
assert (JOBS / running / "deals.sqlite").is_file()
assert web._read_state(done)["state"] == "cleared"
active_stage = JOBS / "staging" / r.get_json()["id"]
# a second init must not touch the upload that is still receiving chunks
assert c.post("/upload/begin", data={"bytes": "10"}).status_code == 200
assert active_stage.is_dir(), "an active upload was deleted"
# once the lock is released (job ended / worker died) it is cleanable
cleanup.release_job_lock(lock)
assert not cleanup.job_locked(JOBS, running)
cleanup.clean(JOBS)
assert not (JOBS / running).exists()

# A download in progress holds the lock too.
dl = fake_job("done")
lk = cleanup.acquire_job_lock(JOBS, dl)
cleanup.clean(JOBS)
assert (JOBS / dl / "hasil.zip").is_file()
cleanup.release_job_lock(lk)

# 2. Missing result -> friendly 410, never 500.
for jid in (dl,):
    cleanup.clean(JOBS)
    for url in (f"/job/{jid}/download", f"/job/{jid}"):
        r = c.get(url)
        body = r.get_data(as_text=True)
        assert r.status_code in (200, 410), (url, r.status_code)
        assert "This result was cleared to free space, please run again." in body, url
# state still 'done' but file gone (e.g. deleted by hand)
gone = fake_job("done")
import shutil  # noqa: E402
shutil.rmtree(JOBS / gone)
r = c.get(f"/job/{gone}/download")
assert r.status_code == 410 and "cleared to free space" in r.get_data(as_text=True)

# 3. Precheck runs AFTER cleanup and sees the freed space.
Usage = namedtuple("Usage", "total used free")
real = chunked.shutil.disk_usage
for st in (JOBS / "staging").iterdir():   # nothing else active in this section
    chunked.shutil.rmtree(st)
big = fake_job("done", size=100 * 1024 * 1024)


def fake_usage(path):
    left = sum(p.stat().st_size for p in JOBS.rglob("*") if p.is_file())
    return Usage(100 * _GIB, 0, 400 * 1024 * 1024 - left)


chunked.shutil.disk_usage = fake_usage
try:
    want = 41 * 1024 * 1024
    assert chunked.cek_ruang(JOBS, want) is not None, "test setup: must not fit before cleanup"
    r = c.post("/upload/begin", data={"bytes": str(want)})
    assert r.status_code == 200, r.get_data(as_text=True)
    assert not (JOBS / big).exists()
    chunked.discard(JOBS, r.get_json()["id"])
    # Still too little with nothing else running -> clear error mentioning cleanup.
    chunked.shutil.disk_usage = lambda p: Usage(100 * _GIB, 0, 1024)
    r = c.post("/upload/begin", data={"bytes": "1000"})
    assert r.status_code == 507, r.status_code
    msg = r.get_json()["error"]
    assert "Not enough free disk space" in msg and "already cleared" in msg, msg
finally:
    chunked.shutil.disk_usage = real

# 4. Queue: another heavy job running -> 202 queued with position, then starts.
heavy = fake_job("running")
hl = cleanup.acquire_job_lock(JOBS, heavy)
r = c.post("/upload/begin", data={"bytes": str(2 * _GIB)})
assert r.status_code == 202, r.status_code
q = r.get_json()
assert q["queued"] and q["position"] == 1
assert "processing another user's file" in q["message"]
r2 = c.post("/upload/begin", data={"bytes": str(3 * _GIB)})
assert r2.status_code == 202 and r2.get_json()["position"] == 2
# poll keeps position
assert c.post("/upload/begin", data={"bytes": str(2 * _GIB), "ticket": q["ticket"]}).get_json()["position"] == 1
# not enough disk for both + other activity -> small upload also queues
chunked.shutil.disk_usage = lambda p: Usage(100 * _GIB, 0, 1024)
try:
    small = c.post("/upload/begin", data={"bytes": "1000"})
    assert small.status_code == 202
    cleanup.queue_leave(JOBS, small.get_json()["ticket"])
finally:
    chunked.shutil.disk_usage = real
# cancel waiting
c.post("/upload/queue/leave", data={"ticket": r2.get_json()["ticket"]})
assert c.post("/upload/begin", data={"bytes": str(2 * _GIB), "ticket": q["ticket"]}).get_json()["position"] == 1
# job ends (lock released) -> the waiting browser's next poll starts the upload
cleanup.release_job_lock(hl)
chunked.shutil.disk_usage = lambda p: Usage(100 * _GIB, 0, 50 * _GIB)
try:
    r = c.post("/upload/begin", data={"bytes": str(2 * _GIB), "ticket": q["ticket"]})
    assert r.status_code == 200, r.get_data(as_text=True)
    assert not (JOBS / heavy).exists(), "finished job not cleaned before start"
    heavy_uid = r.get_json()["id"]
    # the heavy upload itself is now the active heavy work: next heavy one waits
    r = c.post("/upload/begin", data={"bytes": str(2 * _GIB)})
    assert r.status_code == 202
    cleanup.queue_leave(JOBS, r.get_json()["ticket"])
    # 5. stale heartbeat: an upload whose chunks stopped no longer blocks anyone
    st = JOBS / "staging" / heavy_uid
    m = json.loads((st / "manifest.json").read_text())
    m["last_chunk"] = time.time() - cleanup.UPLOAD_IDLE_SECONDS - 5
    (st / "manifest.json").write_text(json.dumps(m))
    r = c.post("/upload/begin", data={"bytes": str(2 * _GIB)})
    assert r.status_code == 200, "stale upload heartbeat still blocked the queue"
    assert not st.exists(), "stale upload not cleaned"
    chunked.discard(JOBS, r.get_json()["id"])
finally:
    chunked.shutil.disk_usage = real

# stale queue ticket expires when the browser stops polling
t = "a" * 32
cleanup.queue_touch(JOBS, t)
p = JOBS / "queue" / t
os.utime(p, (OLD, OLD))
assert cleanup.queue_position(JOBS, "b" * 32) == 1 and not p.exists()

# a job lock held by a dead process is free (kernel drops flock on exit)
import subprocess, sys  # noqa: E402,E401
dead = "d" * 32
subprocess.run([sys.executable, "-c",
                f"import cleanup; cleanup.acquire_job_lock({str(JOBS)!r}, {dead!r})"],
               check=True, cwd=Path(__file__).parent)
assert not cleanup.job_locked(JOBS, dead)

# 6. Failed run deletes its inputs (job dir) and the staged upload.
assert "shutil.rmtree(job, ignore_errors=True)" in Path(web.__file__).read_text()
print("test_cleanup: all checks passed")
