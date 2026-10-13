#!/usr/bin/env python3
"""Free disk at the start of every new run, without touching active work.

What counts as active (safe across the 4 gunicorn workers, which share only
the filesystem):

* A job is active while some process holds an exclusive fcntl.flock on
  JOBS/<job_id>.lock. run() takes the lock before the job directory exists and
  the background thread releases it when the pipeline ends; a download holds it
  while it streams. flock is released by the kernel when a worker dies, so a
  crashed job can never keep its lock (no stale lock possible).
* A chunked upload is active while its manifest heartbeat (last_chunk, updated
  on every chunk) is younger than UPLOAD_IDLE_SECONDS.
* Anything younger than YOUNG_SECONDS is skipped, which covers the instant
  between creating a file and taking its lock.

Everything else in DW_JOB_DIR is old data of finished or abandoned jobs and is
deleted: job directories (inputs, SQLite deal indexes, step files, results),
claimed download directories, idle staging directories, stray temp files and
unheld lock files. Small state JSON files are kept (so status/download pages
can explain what happened) and only removed after STATE_MAX_AGE_SECONDS.
"""
import fcntl
import json
import os
import shutil
import sys
import time
from contextlib import contextmanager
from pathlib import Path


def _env_float(name, default):
    try:
        return float(os.environ.get(name, default))
    except ValueError:
        return float(default)


UPLOAD_IDLE_SECONDS = _env_float("DW_UPLOAD_IDLE_SECONDS", 15 * 60)
YOUNG_SECONDS = 60
STATE_MAX_AGE_SECONDS = 7 * 24 * 3600
QUEUE_TICKET_TTL = 30          # a waiting browser polls every few seconds
HEAVY_BYTES = int(_env_float("DW_HEAVY_BYTES", 1024 ** 3))   # 1 GiB


# --------------------------------------------------------------- job locks
def lock_path(jobs_dir, job_id):
    return Path(jobs_dir) / f"{job_id}.lock"


def acquire_job_lock(jobs_dir, job_id):
    """Open and exclusively lock the job's lock file. Returns the fd or None."""
    fh = open(lock_path(jobs_dir, job_id), "a+b")
    try:
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        fh.close()
        return None
    return fh


def release_job_lock(fh):
    if fh is None:
        return
    try:
        fcntl.flock(fh, fcntl.LOCK_UN)
    finally:
        fh.close()


def job_locked(jobs_dir, job_id):
    """True when another open file (any worker, any thread) holds the job lock."""
    p = lock_path(jobs_dir, job_id)
    try:
        fh = open(p, "rb")
    except OSError:
        return False
    try:
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        return True
    else:
        fcntl.flock(fh, fcntl.LOCK_UN)
        return False
    finally:
        fh.close()


@contextmanager
def admission(jobs_dir):
    """Serialise cleanup + disk check + session creation across workers."""
    with open(Path(jobs_dir) / ".admission.lock", "a+b") as fh:
        fcntl.flock(fh, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(fh, fcntl.LOCK_UN)


# ---------------------------------------------------------------- helpers
def _size(p):
    try:
        if p.is_symlink() or p.is_file():
            return p.lstat().st_size
        total = 0
        for root, _dirs, files in os.walk(p):
            for f in files:
                try:
                    total += os.lstat(os.path.join(root, f)).st_size
                except OSError:
                    pass
        return total
    except OSError:
        return 0


def _remove(p):
    n = _size(p)
    try:
        if p.is_dir() and not p.is_symlink():
            shutil.rmtree(p, ignore_errors=True)
        else:
            p.unlink(missing_ok=True)
    except OSError:
        return 0
    return n if not p.exists() else 0


def _read_json(p):
    try:
        return json.loads(p.read_text())
    except (OSError, ValueError):
        return None


def _write_json(p, data):
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(data))
    tmp.replace(p)


def _mtime(p):
    try:
        return p.stat().st_mtime
    except OSError:
        return 0


def upload_heartbeat(stage):
    m = _read_json(stage / "manifest.json")
    if m and m.get("last_chunk") is not None:
        try:
            return float(m["last_chunk"])
        except (TypeError, ValueError):
            pass
    return _mtime(stage)


def active_uploads(jobs_dir, now=None):
    """[(upload_id, declared_bytes)] for staging sessions with a fresh heartbeat."""
    now = time.time() if now is None else now
    root = Path(jobs_dir) / "staging"
    out = []
    if not root.is_dir():
        return out
    for stage in root.iterdir():
        if stage.is_dir() and now - upload_heartbeat(stage) < UPLOAD_IDLE_SECONDS:
            m = _read_json(stage / "manifest.json") or {}
            out.append((stage.name, int(m.get("declared_bytes") or 0)))
    return out


def active_jobs(jobs_dir):
    """job ids whose lock is currently held."""
    out = []
    for p in Path(jobs_dir).glob("*.lock"):
        if p.name.startswith("."):
            continue
        if job_locked(jobs_dir, p.stem):
            out.append(p.stem)
    return out


def heavy_activity(jobs_dir, keep_upload=None):
    """Other work that uses lots of disk: running segregate jobs or big uploads."""
    jobs = []
    for job_id in active_jobs(jobs_dir):
        info = _read_json(Path(jobs_dir) / f"{job_id}.json") or {}
        if info.get("slug", "segregate") == "segregate" and info.get("state") in ("running", None):
            jobs.append(job_id)
    ups = [u for u, b in active_uploads(jobs_dir) if u != keep_upload and b >= HEAVY_BYTES]
    return jobs + ups


def any_activity(jobs_dir, keep_upload=None):
    ups = [u for u, _b in active_uploads(jobs_dir) if u != keep_upload]
    return active_jobs(jobs_dir) + ups


# ---------------------------------------------------------------- cleanup
def clean(jobs_dir, keep_jobs=(), keep_uploads=(), now=None, log=True):
    """Delete data of every finished/abandoned job. Returns bytes freed."""
    jobs_dir = Path(jobs_dir)
    now = time.time() if now is None else now
    keep_jobs, keep_uploads = set(keep_jobs), set(keep_uploads)
    freed = 0
    removed = []

    stage_root = jobs_dir / "staging"
    if stage_root.is_dir():
        for stage in stage_root.iterdir():
            if stage.name in keep_uploads:
                continue
            if now - upload_heartbeat(stage) < UPLOAD_IDLE_SECONDS:
                continue                      # still receiving chunks
            freed += _remove(stage)
            removed.append(f"staging/{stage.name}")

    for entry in list(jobs_dir.iterdir()):
        name = entry.name
        if name in ("staging", ".admission.lock", "queue"):
            continue
        if now - _mtime(entry) < YOUNG_SECONDS:
            continue                          # created a moment ago, lock may follow
        # .<id>.download  (claimed by a download), <id> dir, <id>.json/.lock/.tmp
        if name.startswith(".") and name.endswith(".download"):
            job_id = name[1:-len(".download")]
        else:
            job_id = name.split(".", 1)[0]
        if job_id in keep_jobs or job_locked(jobs_dir, job_id):
            continue
        if entry.is_dir():
            freed += _remove(entry)
            removed.append(name)
            state_p = jobs_dir / f"{job_id}.json"
            info = _read_json(state_p)
            if info and info.get("state") in ("done", "running"):
                info.update(state="cleared", path=None, cleared=now)
                try:
                    _write_json(state_p, info)
                except OSError:
                    pass
        elif name.endswith(".json"):
            if now - _mtime(entry) > STATE_MAX_AGE_SECONDS:
                freed += _remove(entry)
            else:
                info = _read_json(entry)
                if (info and info.get("state") == "running"
                        and not (jobs_dir / job_id).exists()):
                    info.update(state="cleared", path=None, cleared=now)
                    _write_json(entry, info)
        else:                                  # unheld .lock, stray .tmp, temp files
            freed += _remove(entry)
            removed.append(name)

    if log:
        print(f"cleanup: freed {freed / 1024 ** 3:.2f} GB ({freed} bytes) "
              f"from {len(removed)} item(s) in {jobs_dir}", file=sys.stderr, flush=True)
    return freed


# ------------------------------------------------------------------ queue
def _queue_dir(jobs_dir):
    d = Path(jobs_dir) / "queue"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _valid_ticket(t):
    return bool(t) and len(t) == 32 and all(c in "0123456789abcdef" for c in t)


def queue_touch(jobs_dir, ticket):
    """Create or refresh a waiting ticket. Returns its 1-based position."""
    d = _queue_dir(jobs_dir)
    now = time.time()
    p = d / ticket
    if p.exists():
        os.utime(p, (now, now))
    else:
        p.write_text(str(now))
    return queue_position(jobs_dir, ticket, now)


def queue_position(jobs_dir, ticket, now=None):
    now = time.time() if now is None else now
    d = _queue_dir(jobs_dir)
    live = []
    for p in d.iterdir():
        if now - _mtime(p) > QUEUE_TICKET_TTL:
            p.unlink(missing_ok=True)          # browser stopped polling
            continue
        try:
            created = float(p.read_text() or 0)
        except (OSError, ValueError):
            created = now
        live.append((created, p.name))
    live.sort()
    names = [n for _c, n in live]
    return names.index(ticket) + 1 if ticket in names else len(names) + 1


def queue_leave(jobs_dir, ticket):
    if _valid_ticket(ticket):
        (_queue_dir(jobs_dir) / ticket).unlink(missing_ok=True)
