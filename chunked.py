#!/usr/bin/env python3
"""Chunked upload staging.

Cloudflare rejects any single request body over 100 MB, so a 900 MB batch can
never arrive in one POST. The browser therefore slices each file into pieces and
sends them one at a time; every request stays far below the edge limit while the
total is unlimited. The pieces are appended to a staging file on disk, so the
server never holds a whole upload in memory.

Each file is one staging part. Chunks are appended to that part and then
discarded from memory. There is never a second assembled copy beside the chunks.

Flow:
    GET  /upload/preflight?bytes=&max_part= -> need / free / max_part_fit (auth)
    POST /upload/begin                  -> {"id": "<hex>"}
    POST /upload/chunk  (repeat)        -> {"received": <bytes>}
    POST /upload/ingest                 -> parse one finished part into the
                                          job index, then delete the part file
                                          (Deal Segregator interleaved path)
    GET  /upload/status?id=<hex>        -> chunks + which parts are ingested
    POST /tool/<slug>   upload_id=<hex> -> remaining parts renamed into the job;
                                          if every part is already ingested, only
                                          the session is finalized (zip/output)

A chunk can reach the server while its acknowledgement is lost on the way back
(office firewall, antivirus, Wi-Fi). The browser then sends the same chunk again.
Every committed chunk keeps its size and a short hash in the manifest, so a
repeat of ANY already-stored chunk is acknowledged without writing it twice.

Every /upload/* request is logged to stderr (the pm2 error log) as one line,
so a failed upload leaves evidence on the server. File contents are never logged.
"""
import datetime
import hashlib
import logging
import math
import fcntl
import json
import os
import shutil
import sys
import time
import uuid
from pathlib import Path

from flask import g, jsonify, request
from werkzeug.exceptions import ClientDisconnected
from werkzeug.utils import secure_filename

import cleanup
from upload_limits import MAX_FILES

CHUNK_SUFFIXES = (".csv", ".xlsx", ".xlsm")
# Idle time since the last chunk, not the age of the staging directory.
# With no resume, abandoned staging is discarded promptly (same as cleanup's
# UPLOAD_IDLE_SECONDS, default 15 minutes).
STAGE_MAX_AGE_HOURS = 15 / 60
_GIB = 1024 * 1024 * 1024
# Deal-index overhead measured on this code (journal off, no sqlite temp file,
# workbook deleted as soon as it is inside the zip — the index and the workbook
# are not on disk together). Peak extra disk was the sqlite file in every run:
#   wide 24-column UTF-16, 2.00 GiB, 6,118,728 deals: 97,878,016 bytes (4.56%)
#   wide 24-column UTF-16, 734 MB, 2,095,520 deals: 33,325,056 bytes (4.54%)
#   short 12-column UTF-16, 262 MB, 1,687,662 deals: 26,918,912 bytes (10.27%)
#   short 12-column UTF-8, 168 MB, 2,160,166 deals: 34,447,360 bytes (20.53%)
# The precheck uses the densest of those, rounded up. It does not know how
# wide a row will be. A fixed margin covers the result zip.
OVERHEAD_PER_BYTE = 0.2054
MARGIN_DISK = 256 * 1024 * 1024


def batas_dari_lingkungan(lingkungan=None):
    """Upload ceiling in bytes. DW_MAX_UPLOAD_GB defaults to 12 (binary GB)."""
    sumber = os.environ if lingkungan is None else lingkungan
    mentah = str(sumber.get("DW_MAX_UPLOAD_GB", "12")).strip()
    try:
        gb = float(mentah)
    except ValueError:
        gb = 12.0
    if not math.isfinite(gb) or gb <= 0:
        gb = 12.0
    return int(gb * _GIB)


MAX_TOTAL = batas_dari_lingkungan()


def teks_batas():
    """Short label for the current ceiling, used in errors and on the page."""
    if MAX_TOTAL >= _GIB and MAX_TOTAL % _GIB == 0:
        return f"{MAX_TOTAL // _GIB} GB"
    if MAX_TOTAL >= _GIB:
        return f"{MAX_TOTAL / _GIB:.2f} GB"
    return f"{MAX_TOTAL} bytes"


def ruang_dibutuhkan(ukuran, sudah_tersimpan=False, max_part=None):
    """Peak bytes that must still be free on the job volume.

    With process-and-delete (each part is parsed then unlinked before the next
    is kept), peak disk is the largest part plus the deal index for the whole
    upload (OVERHEAD_PER_BYTE of total size) plus MARGIN_DISK for the result.
    When max_part is omitted it defaults to the total (single-file case), which
    keeps the old S + 20.54%·S + margin bound.

    When the bytes are already on this volume (sudah_tersimpan), only the index
    and margin still have to be free on top of what is stored.
    """
    if ukuran < 0:
        ukuran = 0
    ukuran = int(ukuran)
    if max_part is None:
        max_part = ukuran
    else:
        max_part = max(0, int(max_part))
    ekstra = int(ukuran * OVERHEAD_PER_BYTE) + MARGIN_DISK
    if sudah_tersimpan:
        return ekstra
    return max_part + ekstra



def max_part_fit(ukuran, bebas, freed=0, staged=0):
    """Largest part that still fits: free - index(S) - margin (never negative).

    Used when a single-file upload would exceed free space: the preflight message
    tells the user to split, and quotes this as the max part size that would pass.
    """
    if ukuran < 0:
        ukuran = 0
    ukuran = int(ukuran)
    efektif = max(0, int(bebas) + max(0, int(freed)) - max(0, int(staged)))
    ekstra = int(ukuran * OVERHEAD_PER_BYTE) + MARGIN_DISK
    return max(0, efektif - ekstra)


MSG_SERVER_BUSY = (
    "The server is busy processing another upload right now. "
    "Please wait a few minutes and start the upload again."
)


def pesan_disk(ukuran, bebas, sudah_tersimpan=False, max_part=None, busy=False):
    """English-only disk shortage text. When busy, prefer the wait message."""
    if busy:
        return MSG_SERVER_BUSY
    perlu = ruang_dibutuhkan(ukuran, sudah_tersimpan, max_part=max_part)
    if max_part is None:
        max_part = ukuran
    else:
        max_part = max(0, int(max_part))

    def gb(n):
        return f"{n / _GIB:.1f} GB"

    persen = f"{OVERHEAD_PER_BYTE * 100:.2f}%"
    if sudah_tersimpan:
        return (
            f"Not enough free disk space to process this upload. "
            f"The file is already on the server. "
            f"The deal index needs about {gb(perlu)} more "
            f"({persen} of the total, plus {gb(MARGIN_DISK)} for the result). "
            f"Only {gb(bebas)} is free. "
            f"The upload was not deleted. "
            f"Please wait if another upload is using the server, then start the upload again."
        )
    return (
        f"Not enough free disk space for this upload. "
        f"About {gb(perlu)} free is needed "
        f"(largest part {gb(max_part)} + index ~{persen} of total "
        f"{gb(ukuran)} + {gb(MARGIN_DISK)} margin). "
        f"Only {gb(bebas)} is free. No file was saved. "
        f"Please wait if another upload is using the server, or split into smaller parts, "
        f"then start the upload again."
    )


def teks_dibersihkan(freed):
    return (f" Old data from earlier runs was already cleared before this check "
            f"({freed / _GIB:.1f} GB freed); nothing else can be removed right now.")


def cek_ruang(path, ukuran, sudah_tersimpan=False, freed=None, max_part=None, busy=False):
    """None when the job volume can hold the run, otherwise a message for the user."""
    bebas = shutil.disk_usage(path).free
    if bebas < ruang_dibutuhkan(ukuran, sudah_tersimpan, max_part=max_part):
        pesan = pesan_disk(ukuran, bebas, sudah_tersimpan, max_part=max_part, busy=busy)
        if busy:
            return pesan
        return pesan + (teks_dibersihkan(freed) if freed is not None else "")
    return None


def cek_sebelum_proses(path, ukuran_sudah, ukuran_baru, freed=None, max_part=None, busy=False):
    """Room check once staged bytes are already on this volume.

    ukuran_baru still has to be written (a small form upload). Staged bytes
    do not need to be reserved again: renaming them does not use more space.
    With process-and-delete the peak extra is the index for the whole total
    (conservative); max_part is accepted for message wording only.
    """
    total = max(0, ukuran_sudah) + max(0, ukuran_baru)
    perlu = ruang_dibutuhkan(total, sudah_tersimpan=True, max_part=max_part) + max(0, ukuran_baru)
    bebas = shutil.disk_usage(path).free
    if bebas < perlu:
        pesan = pesan_disk(total, bebas, sudah_tersimpan=ukuran_sudah > 0, max_part=max_part,
                           busy=busy)
        if busy:
            return pesan
        return pesan + (teks_dibersihkan(freed) if freed is not None else "")
    return None


def ringkas_preflight(ukuran, bebas, max_part=None, max_upload=None, freed=0, staged=0, held=0):
    """JSON-ready preflight numbers and English-only message.

    Peak is modelled over the whole interleaved lifecycle (upload one part,
    ingest+delete, repeat, then zip): max_part + index(S) + margin.

    When the peak does not fit:
      * held > 0 and the peak WOULD fit once that other work frees disk -> reason
        "busy" (wait and start again).
      * unsplit file that is too big even for the idle free space -> reason
        "too_big" with a split hint quoting max_part_fit (~8 GB typical).
      * otherwise a plain not-enough message (English only; no VPN advice).
    """
    if max_part is None:
        max_part = ukuran
    max_upload = MAX_TOTAL if max_upload is None else max_upload
    perlu = ruang_dibutuhkan(ukuran, False, max_part=max_part)
    # Count bytes already staged for this check so a leftover upload does not
    # look like free space the new run can spend twice.
    efektif = max(0, bebas + max(0, int(freed)) - max(0, int(staged)))
    held = max(0, int(held))
    ok = ukuran <= max_upload and perlu <= efektif
    fit = max_part_fit(ukuran, bebas, freed=freed, staged=staged)
    idle_free = efektif + held
    fit_idle = max(0, idle_free - int(ukuran * OVERHEAD_PER_BYTE) - MARGIN_DISK)

    def gb(n):
        return round(n / _GIB, 2)

    reason = "ok"
    if ok:
        msg = (
            f"Enough free disk. Need {gb(perlu)} GB (largest part {gb(max_part)} GB + index "
            f"{OVERHEAD_PER_BYTE*100:.2f}% of total {gb(ukuran)} GB + margin "
            f"{gb(MARGIN_DISK)} GB). Free {gb(efektif)} GB. Upload limit "
            f"{gb(max_upload)} GB."
        )
    elif held > 0 and perlu <= idle_free:
        reason = "busy"
        msg = MSG_SERVER_BUSY
    elif max_part >= ukuran and ukuran > 0 and fit_idle > 0 and fit_idle < ukuran:
        reason = "too_big"
        msg = (
            f"Not enough free disk for this file as one upload. "
            f"Need about {gb(perlu)} GB free (largest part {gb(max_part)} GB + index "
            f"{OVERHEAD_PER_BYTE*100:.2f}% of total {gb(ukuran)} GB + margin "
            f"{gb(MARGIN_DISK)} GB). Only {gb(efektif)} GB is free now. "
            f"Split into parts of at most {gb(fit_idle)} GB each, select every part "
            f"together, then start the upload again."
        )
    elif held > 0:
        reason = "busy"
        msg = MSG_SERVER_BUSY
    else:
        reason = "disk"
        msg = (
            f"Not enough free disk. Need {gb(perlu)} GB (largest part {gb(max_part)} GB + index "
            f"{OVERHEAD_PER_BYTE*100:.2f}% of total {gb(ukuran)} GB + margin "
            f"{gb(MARGIN_DISK)} GB). Free {gb(efektif)} GB. Upload limit "
            f"{gb(max_upload)} GB."
        )
        if max_part >= ukuran and ukuran > 0 and fit <= 0:
            msg += (
                " Free space is too small even for a tiny part "
                f"(index + margin alone need "
                f"{gb(int(ukuran * OVERHEAD_PER_BYTE) + MARGIN_DISK)} GB)."
            )
        elif max_part >= ukuran and ukuran > 0 and 0 < fit < ukuran:
            reason = "too_big"
            msg += (
                f" Split into parts of at most {gb(fit)} GB each so the peak fits "
                f"the free space available now."
            )

    return {
        "ok": ok,
        "need": perlu,
        "free": efektif,
        "free_raw": bebas,
        "max_upload": max_upload,
        "max_part": max_part,
        "max_part_fit": fit if reason != "too_big" else max(fit, fit_idle),
        "total": ukuran,
        "freed": max(0, int(freed)),
        "staged": max(0, int(staged)),
        "held": held,
        "reason": reason,
        "message": msg,
        "message_en": msg,
    }

def _storage_error(freed=None, busy=False):
    if busy:
        pesan = MSG_SERVER_BUSY
    else:
        pesan = (
            "Not enough temporary storage on the server for this upload. "
            "No file was saved. "
            "Please wait a few minutes if another upload is using the server, "
            "then start the upload again."
        )
        if freed is not None:
            pesan += teks_dibersihkan(freed)
    body = {"error": pesan, "disk": True, "reason": "busy" if busy else "disk"}
    if freed is not None:
        body["freed"] = freed
    return jsonify(body), 507


def _stage_root(jobs_dir):
    d = jobs_dir / "staging"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _stage_dir(jobs_dir, upload_id):
    """Resolve a client-supplied id to a directory, refusing anything outside.

    The id is generated server-side, but it still arrives from the browser, so
    it is validated as pure hex rather than trusted -- otherwise '../../etc'
    would be a valid directory name.
    """
    if not upload_id or len(upload_id) != 32 or not all(c in "0123456789abcdef" for c in upload_id):
        return None
    d = (_stage_root(jobs_dir) / upload_id).resolve()
    if _stage_root(jobs_dir).resolve() not in d.parents:
        return None
    return d


def _manifest_path(stage):
    return stage / "manifest.json"


def _read_manifest(stage):
    try:
        return json.loads(_manifest_path(stage).read_text())
    except (OSError, ValueError):
        return None


def _write_manifest(stage, data):
    tmp = _manifest_path(stage).with_suffix(".tmp")
    tmp.write_text(json.dumps(data))
    tmp.replace(_manifest_path(stage))


def _aktivitas_unggahan(entry):
    """When the last chunk arrived. Directory age is not progress."""
    manifest = _read_manifest(entry)
    if manifest and manifest.get("last_chunk") is not None:
        try:
            return float(manifest["last_chunk"])
        except (TypeError, ValueError):
            pass
    try:
        return entry.stat().st_mtime
    except OSError:
        return 0


def sweep(jobs_dir, max_age_hours=STAGE_MAX_AGE_HOURS):
    """Delete a staging directory only after it has been idle.

    A multi-gigabyte upload runs for hours. The directory itself is created at
    the start, so its age must not decide. The clock is the last chunk received.
    """
    cutoff = time.time() - max_age_hours * 3600
    root = _stage_root(jobs_dir)
    for entry in root.iterdir():
        try:
            if _aktivitas_unggahan(entry) < cutoff:
                shutil.rmtree(entry, ignore_errors=True)
        except OSError:
            pass


def ingest_dir(stage):
    """Directory holding the Deal Segregator IngestSession for this upload."""
    return Path(stage) / "ingest"


def staged_files(jobs_dir, upload_id):
    """Return finished staged file paths still on disk (not yet ingested).

    Parts already ingested and deleted are omitted. Returns None when the id is
    unknown or the upload is incomplete (bytes short, or a non-ingested part
    missing). An empty list means every part was ingested — the caller should
    finalize the session instead of moving files.
    """
    stage = _stage_dir(jobs_dir, upload_id)
    if not stage or not stage.is_dir():
        return None
    manifest = _read_manifest(stage)
    if not manifest:
        return None
    if manifest["bytes"] != manifest.get("declared_bytes"):
        return None
    if not manifest.get("files"):
        return None
    paths = []
    for entry in manifest["files"]:
        if entry.get("ingested"):
            continue
        p = stage / entry["stored"]
        if not p.is_file():
            return None
        paths.append((p, entry["original"]))
    return paths


def all_ingested(jobs_dir, upload_id):
    """True when every declared part has been ingested into the session."""
    stage = _stage_dir(jobs_dir, upload_id)
    if not stage or not stage.is_dir():
        return False
    manifest = _read_manifest(stage) or {}
    files = manifest.get("files") or []
    if not files:
        return False
    if manifest.get("bytes") != manifest.get("declared_bytes"):
        return False
    return all(e.get("ingested") for e in files)


def discard(jobs_dir, upload_id):
    stage = _stage_dir(jobs_dir, upload_id)
    if stage:
        shutil.rmtree(stage, ignore_errors=True)


# ------------------------------------------------------------------ request log
log = logging.getLogger("dw.upload")
if not log.handlers:
    _h = logging.StreamHandler(sys.stderr)
    _h.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(_h)
    log.setLevel(logging.INFO)
    log.propagate = False


def _ip():
    """Client IP as Cloudflare saw it. cloudflared connects from 127.0.0.1."""
    return (request.headers.get("CF-Connecting-IP")
            or (request.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
            or request.remote_addr or "-")


def _bidang(nama):
    """A form/query field, without forcing a body parse the view did not do."""
    nilai = request.args.get(nama)
    if nilai is None and "form" in request.__dict__:
        nilai = request.form.get(nama)
    return nilai


def _catat(status, catatan=None):
    mulai = g.get("upload_t0")
    ms = int((time.perf_counter() - mulai) * 1000) if mulai else -1
    user = getattr(request, "user", None)
    nama = user["username"] if user else "-"
    bagian = [
        "upload",
        datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        request.method, request.path, str(status),
        f"bytes={request.content_length if request.content_length is not None else '-'}",
        f"ms={ms}",
        f"id={(_bidang('id') or '-')[:32]}",
        f"file={_bidang('index') or '-'}",
        f"seq={_bidang('seq') or '-'}",
        f"ip={_ip()}",
        f"user={nama}",
    ]
    catatan = catatan or g.get("upload_note")
    if catatan:
        bagian.append(f"note={str(catatan)[:160]}")
    log.info(" ".join(bagian))


def _username():
    user = getattr(request, "user", None)
    try:
        return (user["username"] if user else None) or None
    except (KeyError, TypeError, IndexError):
        return None


def _jam(ts):
    try:
        return datetime.datetime.fromtimestamp(float(ts)).strftime("%H:%M")
    except (TypeError, ValueError, OSError):
        return None


_NAMA_TOOL = {"segregate": "Deal Segregator", "dw": "Generate D&W", "pl-desk": "PL by Desk",
              "segregated": "Segregated"}


def _apa(b):
    """Short description of what a blocker is doing. No file names of other users."""
    if b["kind"] == "job":
        tool = _NAMA_TOOL.get(b.get("slug") or "", "a run")
        jam = _jam(b.get("started"))
        return f"{tool} run, processing" + (f" since {jam}" if jam else "")
    gb = (b.get("bytes") or 0) / _GIB
    return f"uploading {gb:.1f} GB" if gb >= 0.1 else "uploading"


def pesan_antrean(blockers, user, ahead=()):
    """(message, blocked_by, job_url) for a 202 queued answer.

    Names the account the server is busy with. Legacy entries without an owner
    get the generic text. The requester's own running job gets its own text
    and a link: it is not cancelled (see upload_begin).
    """
    milik = [b for b in blockers if user and b.get("user") == user]
    sendiri = [b for b in milik if b["kind"] == "job" and b.get("state") in ("running", None)]
    lain = [b for b in blockers if not (user and b.get("user") == user)]
    if milik and not sendiri and not lain:
        return ("Your earlier result is still being downloaded. This upload starts "
                "automatically when that finishes.",
                [{"kind": b["kind"], "user": b.get("user"), "own": True} for b in milik], None)
    blocked_by = [{"kind": b["kind"], "user": b.get("user"), "own": bool(user and b.get("user") == user)}
                  for b in blockers]
    job_url = None
    if sendiri:
        b = sendiri[0]
        job_url = (f"/kvb/job/{b['id']}" if b.get("company") == "kvb" else f"/job/{b['id']}")
        nama = f" ({b['name']})" if b.get("name") else ""
        jam = _jam(b.get("started"))
        pesan = (f"Your previous run{nama} is still processing"
                 + (f" (started {jam})" if jam else "") + ". "
                 "A running job cannot be stopped safely, so this upload waits and starts "
                 "automatically when it finishes. You can follow it at " + job_url + ".")
        if lain:
            pesan += " " + pesan_antrean(lain, user)[0]
        return pesan, blocked_by, job_url
    if lain:
        nama = sorted({b["user"] for b in lain if b.get("user")})
        apa = "; ".join(sorted({(f"{b['user']}: " if b.get('user') else "") + _apa(b) for b in lain}))
        if nama:
            pesan = (f"The server is currently processing another user's file ({apa}). "
                     "Your run is queued and will start automatically when it finishes.")
        else:
            pesan = ("The server is currently processing another user's file. Your run is "
                     "queued and will start automatically when it finishes.")
        return pesan, blocked_by, job_url
    depan = sorted({u for u in ahead if u and u != user})
    if depan:
        return (f"The server is currently processing another user's queued upload "
                f"({', '.join(depan)}). Yours starts automatically after it.",
                blocked_by, job_url)
    return ("The server is currently processing another user's file. Your run is queued and "
            "will start automatically when it finishes.", blocked_by, job_url)


def _hash(data):
    return hashlib.sha256(data).hexdigest()[:16]


def _tidak_lengkap(alasan):
    """The body arrived short: the connection dropped mid-request. Retryable."""
    g.upload_note = alasan
    return jsonify({"error": "The chunk arrived incomplete (the connection dropped while "
                             "it was being sent). It will be sent again.",
                    "retry": True}), 408


def _sudah_ada(entry, seq, data, sidik, target):
    """True when chunk `seq` of this file is already stored with these exact bytes."""
    sizes, hashes = entry.get("sizes"), entry.get("hashes")
    if (isinstance(sizes, list) and isinstance(hashes, list)
            and len(sizes) == len(hashes) == entry["chunks"]):
        return sizes[seq] == len(data) and hashes[seq] == sidik
    # Manifest written before hashes were kept: only the last chunk can be checked.
    if seq != entry["chunks"] - 1 or not target.is_file() or target.stat().st_size < len(data):
        return False
    with open(target, "rb") as fh:
        fh.seek(-len(data), 2)
        return fh.read() == data



def bytes_held_by_others(jobs_dir, user=None, keep_upload=None):
    """Disk bytes held by other users' active staging and locked jobs.

    Used to tell 'server busy, wait' apart from 'this file is too big even when
    the server is idle'. The caller's own staging is counted separately as
    `staged` and is not included here.
    """
    held = 0
    root = Path(jobs_dir) / "staging"
    if root.is_dir():
        for stage in root.iterdir():
            if not stage.is_dir() or stage.name == keep_upload:
                continue
            m = _read_manifest(stage) or {}
            owner = m.get("user")
            if user and owner == user:
                continue
            try:
                held += sum(f.stat().st_size for f in stage.rglob("*") if f.is_file())
            except OSError:
                pass
    for job_id in cleanup.active_jobs(jobs_dir):
        info = cleanup._read_json(Path(jobs_dir) / f"{job_id}.json") or {}
        if user and info.get("user") == user:
            continue
        held += cleanup._size(Path(jobs_dir) / job_id)
    return held


def register(app, jobs_dir):
    """Attach the upload endpoints. Auth is handled by the app's global hook."""

    @app.before_request
    def _upload_mulai():
        if request.path.startswith("/upload/"):
            g.upload_t0 = time.perf_counter()

    @app.after_request
    def _upload_log(resp):
        if request.path.startswith("/upload/"):
            g.upload_logged = True
            try:
                _catat(resp.status_code)
            except Exception:                    # logging must never break an upload
                pass
        return resp

    @app.teardown_request
    def _upload_log_error(exc=None):
        # after_request does not run when the view raised: log that too.
        if request.path.startswith("/upload/") and not g.get("upload_logged"):
            try:
                _catat(500, f"exception {type(exc).__name__}: {exc}" if exc else "no response")
            except Exception:
                pass

    @app.get("/upload/status")
    def upload_status():
        """What the server already holds, so a browser can resume after a lost ack."""
        upload_id = request.args.get("id")
        stage = _stage_dir(jobs_dir, upload_id)
        manifest = _read_manifest(stage) if stage and stage.is_dir() else None
        if manifest is None:
            return jsonify({"error": "This upload has expired. Please pick the files again."}), 404
        return jsonify({"id": upload_id, "bytes": manifest["bytes"],
                        "declared_bytes": manifest.get("declared_bytes"),
                        "files": [{"index": e["index"], "chunks": e["chunks"],
                                   "bytes": e["bytes"],
                                   "name": e.get("original"),
                                   "hashes": e.get("hashes"),
                                   "ingested": bool(e.get("ingested")),
                                   "kind": e.get("ingest_kind")}
                                  for e in manifest["files"]],
                        "ingested": [e["index"] for e in manifest["files"]
                                     if e.get("ingested")]})

    @app.get("/upload/preflight")
    def upload_preflight():
        """Cheap disk check before any chunk is sent. CSRF-safe GET, auth required."""
        try:
            requested = int(request.args.get("bytes", "0"))
        except ValueError:
            return jsonify({"error": "Invalid upload size.", "ok": False}), 400
        try:
            max_part = request.args.get("max_part")
            max_part = int(max_part) if max_part not in (None, "") else requested
        except ValueError:
            return jsonify({"error": "Invalid max_part.", "ok": False}), 400
        if requested < 0 or max_part < 0:
            return jsonify({"error": "Invalid upload size.", "ok": False}), 400
        if requested > MAX_TOTAL:
            info = ringkas_preflight(requested, 0, max_part=max_part)
            info["ok"] = False
            info["error"] = f"The selected files exceed the {teks_batas()} limit."
            info["message"] = info["error"]
            return jsonify(info), 413
        user = _username()
        # Count this user's currently staged bytes so a leftover upload is not
        # treated as free space a second time. Preflight does not delete them.
        staged = 0
        root = _stage_root(jobs_dir)
        if root.is_dir() and user:
            for stage in root.iterdir():
                if not stage.is_dir():
                    continue
                m = _read_manifest(stage) or {}
                if m.get("user") == user:
                    try:
                        staged += sum(f.stat().st_size for f in stage.rglob("*") if f.is_file())
                    except OSError:
                        pass
        bebas = shutil.disk_usage(_stage_root(jobs_dir)).free
        held = bytes_held_by_others(jobs_dir, user=user)
        # Also treat live blockers (fresh heartbeat / locked job) as "busy" even
        # when their on-disk size is still small — another upload is using the server.
        if held <= 0 and cleanup.all_blockers(jobs_dir):
            held = max(held, 1)
        info = ringkas_preflight(requested, bebas, max_part=max_part, staged=staged, held=held)
        return jsonify(info), (200 if info["ok"] else 507)

    @app.post("/upload/begin")
    def upload_begin():
        try:
            requested = int(request.form.get("bytes", "0"))
        except ValueError:
            return jsonify({"error": "Invalid upload size."}), 400
        try:
            raw_max = request.form.get("max_part")
            max_part = int(raw_max) if raw_max not in (None, "") else requested
        except ValueError:
            return jsonify({"error": "Invalid max_part."}), 400
        if max_part < 0:
            max_part = requested
        if requested < 0 or requested > MAX_TOTAL:
            return jsonify({"error": f"The selected files exceed the {teks_batas()} limit."}), 413
        ticket = (request.form.get("ticket") or "").strip()
        user = _username()
        if cleanup.ticket_superseded(jobs_dir, ticket):
            # This tab waited in the queue, then the same account started a newer upload.
            return jsonify({"error": "This upload was replaced by a newer upload from your "
                                     "account (another tab or a new attempt). Nothing was "
                                     "uploaded here.", "superseded": True}), 409
        if not cleanup._valid_ticket(ticket):
            ticket = uuid.uuid4().hex
        # One worker at a time: clean, measure, then reserve. Without this two
        # workers could both see the same free space and both start.
        with cleanup.admission(jobs_dir):
            # The requester's own older upload (stalled, abandoned, or another tab)
            # and their older queue slot never make them wait: the newest begin wins.
            # Before this, those counted as "another user's" heavy upload for up to
            # 15 minutes, so a user queued behind themselves.
            diganti, bebas_lama = cleanup.supersede_user_uploads(jobs_dir, user)
            tiket_lama = cleanup.queue_drop_user(jobs_dir, user, keep=ticket)
            if diganti or tiket_lama:
                g.upload_note = (f"superseded own uploads {len(diganti)}, "
                                 f"own queue tickets {len(tiket_lama)}")
            # sweep() is the idle staging cleaner that used to sit unused; clean()
            # already removes idle staging by heartbeat, and sweep covers the same
            # directories by last-chunk age so both stay consistent.
            sweep(jobs_dir)
            freed = cleanup.clean(jobs_dir) + bebas_lama
            salah = cek_ruang(_stage_root(jobs_dir), requested, sudah_tersimpan=False,
                              freed=freed, max_part=max_part)
            berat = requested >= cleanup.HEAVY_BYTES
            blockers = cleanup.heavy_blockers(jobs_dir) if berat else []
            if salah and not blockers:
                blockers = cleanup.all_blockers(jobs_dir)
            sibuk = bool(blockers) if (berat or salah) else False
            posisi = cleanup.queue_position(jobs_dir, ticket)
            if sibuk or (berat and posisi > 1):
                posisi = cleanup.queue_touch(jobs_dir, ticket, user)
                pesan, blocked_by, job_url = pesan_antrean(
                    blockers, user, cleanup.queue_ahead(jobs_dir, ticket))
                g.upload_note = "queued behind " + (", ".join(
                    sorted({(b["user"] or "?") + ":" + b["kind"] for b in blocked_by})) or "queue")
                return jsonify({
                    "queued": True, "ticket": ticket, "position": posisi,
                    "message": pesan, "blocked_by": blocked_by, "job_url": job_url,
                }), 202
            cleanup.queue_leave(jobs_dir, ticket)
            if salah:
                bebas = shutil.disk_usage(_stage_root(jobs_dir)).free
                held = bytes_held_by_others(jobs_dir, user=user)
                if held <= 0 and blockers:
                    held = max(held, 1)
                info = ringkas_preflight(requested, bebas, max_part=max_part, freed=freed, held=held)
                err = info["message"] if info.get("reason") in ("busy", "too_big") else salah
                # Prefer English classified message; strip bilingual leftovers from cek_ruang.
                return jsonify({"error": err, "freed": freed, **info}), 507
            upload_id = uuid.uuid4().hex
            stage = _stage_root(jobs_dir) / upload_id
            try:
                sekarang = time.time()
                stage.mkdir()
                _write_manifest(stage, {"files": [], "bytes": 0, "user": user,
                                        "declared_bytes": requested,
                                        "max_part": max_part,
                                        "started": sekarang,
                                        "last_chunk": sekarang})
            except OSError:
                shutil.rmtree(stage, ignore_errors=True)
                busy = bool(cleanup.all_blockers(jobs_dir))
                return _storage_error(busy=busy)
        return jsonify({"id": upload_id, "freed": freed})

    @app.post("/upload/queue/leave")
    def upload_queue_leave():
        cleanup.queue_leave(jobs_dir, (request.form.get("ticket") or "").strip())
        return jsonify({"left": True})

    @app.post("/upload/discard")
    def upload_discard():
        """Discard the caller's own staged upload (retries exhausted, cancel, tab closed).
        Another account's upload is never removed; legacy uploads without an owner are."""
        upload_id = request.form.get("id")
        stage = _stage_dir(jobs_dir, upload_id)
        manifest = _read_manifest(stage) if stage and stage.is_dir() else None
        owner = (manifest or {}).get("user")
        if manifest is not None and owner and owner != _username():
            g.upload_note = "discard refused: upload belongs to another account"
            return jsonify({"discarded": False}), 403
        g.upload_note = f"discard ({request.form.get('reason') or 'client'})"
        discard(jobs_dir, upload_id)
        return jsonify({"discarded": True})


    @app.post("/upload/ingest")
    def upload_ingest():
        """Parse one finished staged part into the Deal Segregator session, then delete it.

        Sequential with the browser: upload part i fully, call this, then upload
        part i+1. Peak disk is therefore ~max(part) + index + margin, not the sum
        of every part. FX / account-type files should be ingested before deals.
        """
        import deal_segregator as ds
        upload_id = (request.form.get("id") or "").strip()
        stage = _stage_dir(jobs_dir, upload_id)
        if stage and not stage.is_dir() and cleanup.superseded(stage):
            return jsonify({"error": "This upload was replaced by a newer upload from your "
                                     "account (another tab or a new attempt).",
                            "superseded": True}), 410
        if not stage or not stage.is_dir():
            return jsonify({"error": "This upload has expired. Please pick the files again."}), 404
        try:
            index = int(request.form.get("index", ""))
        except ValueError:
            return jsonify({"error": "Malformed ingest request."}), 400
        allow_mt4 = (request.form.get("allow_mt4") or "").strip() in ("1", "true", "yes")

        with open(stage / ".lock", "a+b") as lock:
            fcntl.flock(lock, fcntl.LOCK_EX)
            manifest = _read_manifest(stage)
            if manifest is None:
                return jsonify({"error": "This upload is no longer valid. Please start again."}), 404
            owner = manifest.get("user")
            if owner and owner != _username():
                return jsonify({"error": "This upload belongs to another account."}), 403
            entries = {e["index"]: e for e in manifest["files"]}
            entry = entries.get(index)
            if entry is None:
                return jsonify({"error": "That file has not been uploaded yet."}), 404
            if entry.get("ingested"):
                g.upload_note = f"ingest duplicate index={index}"
                return jsonify({"ingested": True, "duplicate": True,
                                "index": index, "name": entry.get("original"),
                                "kind": entry.get("ingest_kind")})
            target = stage / entry["stored"]
            if not target.is_file():
                return jsonify({"error": "The staged file is missing. Please start again."}), 404
            # Part must be complete relative to its own bytes so far; the browser
            # only calls ingest after the last chunk of that file has been acked.
            state = ingest_dir(stage)
            state.mkdir(parents=True, exist_ok=True)
            try:
                sess = ds.IngestSession.load(state, allow_mt4=allow_mt4)
                rec = sess.ingest(target, hapus=True, original_name=entry.get("original"))
            except SystemExit as exc:
                pesan = str(exc) or "STOP: ingest failed."
                g.upload_note = f"ingest STOP index={index}"
                return jsonify({"error": pesan, "stop": True}), 400
            except Exception as exc:
                g.upload_note = f"ingest exception index={index}: {type(exc).__name__}"
                return jsonify({"error": f"STOP: failed while ingesting "
                                         f"{entry.get('original')}: {exc}"}), 500
            entry["ingested"] = True
            entry["ingest_kind"] = rec.get("kind")
            entry["ingested_at"] = time.time()
            # Part file is gone; keep manifest bytes as the cumulative received
            # total (for resume / declared_bytes check) but do not count the
            # deleted file toward disk.
            manifest["last_chunk"] = time.time()
            _write_manifest(stage, manifest)
            g.upload_note = f"ingested index={index} kind={rec.get('kind')}"
            return jsonify({"ingested": True, "index": index,
                            "name": entry.get("original"), "kind": rec.get("kind"),
                            "ingested_indexes": [e["index"] for e in manifest["files"]
                                                 if e.get("ingested")]})

    @app.post("/upload/chunk")
    def upload_chunk():
        try:
            form = request.form
            blob = request.files.get("chunk")
        except ClientDisconnected:
            return _tidak_lengkap("client disconnected while sending the body")
        if (request.content_length or 0) > 0 and (not form.get("id") or not blob):
            # A body cut off mid-way parses as an empty form; do not call it expired.
            return _tidak_lengkap(f"incomplete body: Content-Length {request.content_length}, "
                                  f"fields {sorted(form.keys())}, chunk {'yes' if blob else 'no'}")
        stage = _stage_dir(jobs_dir, request.form.get("id"))
        if stage and not stage.is_dir() and cleanup.superseded(stage):
            g.upload_note = "chunk for an upload superseded by a newer one"
            return jsonify({"error": "This upload was replaced by a newer upload from your "
                                     "account (another tab or a new attempt).",
                            "superseded": True}), 410
        if not stage or not stage.is_dir():
            return jsonify({"error": "This upload has expired. Please pick the files again."}), 404
        original = (request.form.get("name") or "").strip()
        if not original.lower().endswith(CHUNK_SUFFIXES):
            return jsonify({"error": f"'{original}' is not a .csv, .xlsx or .xlsm file."}), 400
        try:
            index = int(request.form.get("index", ""))
            seq = int(request.form.get("seq", ""))
        except ValueError:
            return jsonify({"error": "Malformed upload request."}), 400
        if not 0 <= index < MAX_FILES:
            return jsonify({"error": f"At most {MAX_FILES} files can be sent in one run."}), 400

        if not blob:
            return jsonify({"error": "Malformed upload request."}), 400
        try:
            data = blob.read()
        except ClientDisconnected:
            return _tidak_lengkap("client disconnected while reading the chunk")
        size = request.form.get("size")
        if size not in (None, "") and size.isdigit() and int(size) != len(data):
            return _tidak_lengkap(f"short read: {len(data)} of {size} bytes")
        sidik = _hash(data)

        # Four Gunicorn workers can receive a retry and its original request at
        # once. Lock the upload while checking sequence, appending, and updating
        # the manifest so duplicate bytes can never be committed concurrently.
        with open(stage / ".lock", "a+b") as lock:
            fcntl.flock(lock, fcntl.LOCK_EX)
            manifest = _read_manifest(stage)
            if manifest is None:
                return jsonify({"error": "This upload is no longer valid. Please start again."}), 404
            safe_suffix = Path(secure_filename(original) or "upload").suffix.lower()
            stored = f"part-{index:02d}{safe_suffix or '.csv'}"
            target = stage / stored
            entries = {e["index"]: e for e in manifest["files"]}
            entry = entries.get(index)
            expected = entry["chunks"] if entry else 0
            if 0 <= seq < expected:
                # Already stored: the ack was lost and the browser sent it again.
                if _sudah_ada(entry, seq, data, sidik, target):
                    manifest["last_chunk"] = time.time()
                    _write_manifest(stage, manifest)
                    g.upload_note = "duplicate chunk, already stored"
                    return jsonify({"received": manifest["bytes"], "duplicate": True})
                g.upload_note = f"repeat of chunk {seq} with different bytes"
                return jsonify({"error": "Chunk arrived out of order.", "expected": expected}), 409
            if seq != expected:
                g.upload_note = f"out of order: got {seq}, expected {expected}"
                return jsonify({"error": "Chunk arrived out of order.", "expected": expected}), 409
            if manifest["bytes"] + len(data) > manifest.get("declared_bytes", 0):
                shutil.rmtree(stage, ignore_errors=True)
                return jsonify({"error": "Received bytes exceed the size declared when the upload started. "
                                         "Please choose the files again."}), 413
            if manifest["bytes"] + len(data) > MAX_TOTAL:
                shutil.rmtree(stage, ignore_errors=True)
                return jsonify({"error": f"The files add up to more than {teks_batas()}. "
                                         "Please split them into several smaller runs."}), 413

            try:
                with open(target, "ab") as fh:
                    fh.write(data)
                if entry:
                    entry["chunks"] += 1
                    entry["bytes"] += len(data)
                    entry.setdefault("sizes", []).append(len(data))
                    entry.setdefault("hashes", []).append(sidik)
                else:
                    manifest["files"].append({"index": index, "original": original,
                                              "stored": stored, "chunks": 1, "bytes": len(data),
                                              "sizes": [len(data)], "hashes": [sidik]})
                manifest["files"].sort(key=lambda e: e["index"])
                manifest["bytes"] += len(data)
                manifest["last_chunk"] = time.time()
                _write_manifest(stage, manifest)
            except OSError:
                shutil.rmtree(stage, ignore_errors=True)
                busy = bool(cleanup.all_blockers(jobs_dir))
                return _storage_error(busy=busy)
            return jsonify({"received": manifest["bytes"]})
