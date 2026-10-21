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

Ownership (who is the server busy with): /upload/begin writes the username into
the staging manifest, the run writes it into the job state JSON, and a queue
ticket stores it too. A user's own older upload and queue slot are superseded by
their new upload instead of queueing them behind themselves. Other users' data
is never removed here; it only expires through the idle rules above.

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
# A big upload makes another big upload wait only while its chunks keep coming.
# The browser's longest normal gap is one 180 s request timeout plus a 30 s
# backoff, so 5 minutes of silence means its owner has gone away. Its staging
# directory is still only deleted after UPLOAD_IDLE_SECONDS (disk is checked
# separately, so its bytes are still counted until then).
UPLOAD_BLOCK_SECONDS = _env_float("DW_UPLOAD_BLOCK_SECONDS", 5 * 60)
SUPERSEDED_TTL = 15 * 60       # how long a replaced upload/ticket can say so
HEAVY_BYTES = int(_env_float("DW_HEAVY_BYTES", 1024 ** 3))   # 1 GiB


# --------------------------------------------------------------- job locks
def lock_path(jobs_dir, job_id):
    return Path(jobs_dir) / f"{job_id}.lock"


def acquire_job_lock(jobs_dir, job_id):
    """Open and exclusively lock the job's lock file. Returns the fd or None.

    release_job_lock() unlinks the file, so after locking we confirm the path
    still names the inode we locked; if a releaser unlinked it in between, we
    hold a lock on an orphan and must retry on the fresh file.
    """
    path = lock_path(jobs_dir, job_id)
    for _ in range(5):
        fh = open(path, "a+b")
        try:
            fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            fh.close()
            return None
        try:
            same = os.stat(path).st_ino == os.fstat(fh.fileno()).st_ino
        except FileNotFoundError:
            same = False
        if same:
            return fh
        fcntl.flock(fh, fcntl.LOCK_UN)
        fh.close()
    return None


def release_job_lock(fh):
    """Unlock and remove the lock file, so a finished or refused run leaves no
    <job_id>.lock behind in JOBS. Unlinking happens while still holding the
    lock; acquire_job_lock() detects and retries on the orphaned inode."""
    if fh is None:
        return
    try:
        try:
            if os.stat(fh.name).st_ino == os.fstat(fh.fileno()).st_ino:
                os.unlink(fh.name)
        except OSError:
            pass
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


def active_uploads(jobs_dir, now=None, idle=None):
    """[(upload_id, declared_bytes)] for staging sessions with a fresh heartbeat."""
    now = time.time() if now is None else now
    idle = UPLOAD_IDLE_SECONDS if idle is None else idle
    root = Path(jobs_dir) / "staging"
    out = []
    if not root.is_dir():
        return out
    for stage in root.iterdir():
        if stage.is_dir() and now - upload_heartbeat(stage) < idle:
            m = _read_json(stage / "manifest.json") or {}
            out.append((stage.name, int(m.get("declared_bytes") or 0)))
    return out


def upload_owner(jobs_dir, upload_id):
    m = _read_json(Path(jobs_dir) / "staging" / upload_id / "manifest.json") or {}
    return m.get("user") or None


def describe_activity(jobs_dir, kind, ident):
    """{kind, id, user, ...} for one active job or upload. No paths, no secrets."""
    jobs_dir = Path(jobs_dir)
    if kind == "job":
        info = _read_json(jobs_dir / f"{ident}.json") or {}
        return {"kind": "job", "id": ident, "user": info.get("user") or None,
                "slug": info.get("slug"), "company": info.get("company"),
                "state": info.get("state"), "name": info.get("name"),
                "started": info.get("started")}
    m = _read_json(jobs_dir / "staging" / ident / "manifest.json") or {}
    return {"kind": "upload", "id": ident, "user": m.get("user") or None,
            "bytes": int(m.get("declared_bytes") or 0), "started": m.get("started")}


def active_jobs(jobs_dir):
    """job ids whose lock is currently held."""
    out = []
    for p in Path(jobs_dir).glob("*.lock"):
        if p.name.startswith("."):
            continue
        if job_locked(jobs_dir, p.stem):
            out.append(p.stem)
    return out


def heavy_blockers(jobs_dir, keep_upload=None):
    """Work that uses lots of disk: running segregate jobs, or big uploads whose
    chunks arrived within UPLOAD_BLOCK_SECONDS. -> [describe_activity(...)]"""
    out = []
    for job_id in active_jobs(jobs_dir):
        info = _read_json(Path(jobs_dir) / f"{job_id}.json") or {}
        if info.get("slug", "segregate") == "segregate" and info.get("state") in ("running", None):
            out.append(describe_activity(jobs_dir, "job", job_id))
    for u, b in active_uploads(jobs_dir, idle=UPLOAD_BLOCK_SECONDS):
        if u != keep_upload and b >= HEAVY_BYTES:
            out.append(describe_activity(jobs_dir, "upload", u))
    return out


def all_blockers(jobs_dir, keep_upload=None):
    """Every active job (lock held) and every upload with a live heartbeat."""
    out = [describe_activity(jobs_dir, "job", j) for j in active_jobs(jobs_dir)]
    out += [describe_activity(jobs_dir, "upload", u)
            for u, _b in active_uploads(jobs_dir) if u != keep_upload]
    return out


def heavy_activity(jobs_dir, keep_upload=None):
    return [b["id"] for b in heavy_blockers(jobs_dir, keep_upload)]


def any_activity(jobs_dir, keep_upload=None):
    return [b["id"] for b in all_blockers(jobs_dir, keep_upload)]


# ------------------------------------------------------------ supersede (same user)
def _tanda_diganti(path):
    try:
        path.write_text(str(time.time()))
    except OSError:
        pass


def superseded(path):
    """True when `path` (a staging id or queue ticket marker) was replaced recently."""
    p = Path(str(path) + ".superseded")
    return p.is_file() and time.time() - _mtime(p) < SUPERSEDED_TTL


def supersede_user_uploads(jobs_dir, user, keep=()):
    """Discard this user's own older staged uploads. -> ([upload_id], bytes freed).

    Only staging directories whose manifest names `user` are touched; other
    users' uploads and legacy ones without an owner are left alone. The stage's
    chunk lock is taken first, so a chunk being appended right now finishes
    before the directory goes. A marker lets that older tab say it was replaced.
    """
    if not user:
        return [], 0
    root = Path(jobs_dir) / "staging"
    if not root.is_dir():
        return [], 0
    keep = set(keep)
    gone = []
    freed = 0
    for stage in root.iterdir():
        if not stage.is_dir() or stage.name in keep:
            continue
        m = _read_json(stage / "manifest.json") or {}
        if m.get("user") != user:
            continue
        ukuran = _size(stage)
        try:
            with open(stage / ".lock", "a+b") as lock:
                fcntl.flock(lock, fcntl.LOCK_EX)
                shutil.rmtree(stage, ignore_errors=True)
        except OSError:
            shutil.rmtree(stage, ignore_errors=True)
        if not stage.exists():
            freed += ukuran
        _tanda_diganti(root / f"{stage.name}.superseded")
        gone.append(stage.name)
    return gone, freed


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


def queue_touch(jobs_dir, ticket, user=None):
    """Create or refresh a waiting ticket. Returns its 1-based position."""
    d = _queue_dir(jobs_dir)
    now = time.time()
    p = d / ticket
    if p.exists():
        os.utime(p, (now, now))
    else:
        p.write_text(json.dumps({"created": now, "user": user}))
    return queue_position(jobs_dir, ticket, now)


def _baca_tiket(p, now):
    """(created, user). Old tickets hold only the creation time as text."""
    try:
        teks = p.read_text()
    except OSError:
        return now, None
    try:
        data = json.loads(teks or "0")
    except ValueError:
        return now, None
    if isinstance(data, dict):
        try:
            return float(data.get("created") or now), data.get("user") or None
        except (TypeError, ValueError):
            return now, data.get("user") or None
    try:
        return float(data), None
    except (TypeError, ValueError):
        return now, None


def queue_entries(jobs_dir, now=None):
    """[(created, ticket, user)] of live tickets, oldest first. Expired ones,
    and superseded markers past their TTL, are removed here."""
    now = time.time() if now is None else now
    d = _queue_dir(jobs_dir)
    live = []
    for p in d.iterdir():
        if p.name.endswith(".superseded"):
            if now - _mtime(p) > SUPERSEDED_TTL:
                p.unlink(missing_ok=True)
            continue
        if now - _mtime(p) > QUEUE_TICKET_TTL:
            p.unlink(missing_ok=True)          # browser stopped polling
            continue
        created, user = _baca_tiket(p, now)
        live.append((created, p.name, user))
    live.sort()
    return live


def queue_position(jobs_dir, ticket, now=None):
    names = [n for _c, n, _u in queue_entries(jobs_dir, now)]
    return names.index(ticket) + 1 if ticket in names else len(names) + 1


def queue_ahead(jobs_dir, ticket, now=None):
    """Users of the live tickets ahead of `ticket` (None for unknown owners)."""
    out = []
    for _c, name, user in queue_entries(jobs_dir, now):
        if name == ticket:
            break
        out.append(user)
    return out


def queue_drop_user(jobs_dir, user, keep=None):
    """Remove this user's other waiting tickets (an older tab or attempt).
    Each leaves a marker so that tab's next poll is told it was replaced."""
    if not user:
        return []
    d = _queue_dir(jobs_dir)
    gone = []
    for _c, name, owner in queue_entries(jobs_dir):
        if owner == user and name != keep:
            (d / name).unlink(missing_ok=True)
            _tanda_diganti(d / f"{name}.superseded")
            gone.append(name)
    return gone


def ticket_superseded(jobs_dir, ticket):
    return _valid_ticket(ticket) and superseded(_queue_dir(jobs_dir) / ticket)


def queue_leave(jobs_dir, ticket):
    if _valid_ticket(ticket):
        (_queue_dir(jobs_dir) / ticket).unlink(missing_ok=True)
