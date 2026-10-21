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
    POST /upload/begin                  -> {"id": "<hex>"}
    POST /upload/chunk  (repeat)        -> {"received": <bytes>}
    GET  /upload/status?id=<hex>        -> what the server already has (resume)
    POST /tool/<slug>   upload_id=<hex> -> the part is renamed into the job

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
# A 10 GB upload can take several hours as long as chunks keep arriving.
STAGE_MAX_AGE_HOURS = 1
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


def ruang_dibutuhkan(ukuran, sudah_tersimpan=False):
    """Bytes that must still be free on the job volume.

    The upload is stored once. Beside it the deal index needs OVERHEAD_PER_BYTE
    of the file, plus MARGIN_DISK for the result. When the file is already on
    this volume, only that extra space has to be free.
    """
    if ukuran < 0:
        ukuran = 0
    ekstra = int(ukuran * OVERHEAD_PER_BYTE) + MARGIN_DISK
    if sudah_tersimpan:
        return ekstra
    return int(ukuran) + ekstra


def pesan_disk(ukuran, bebas, sudah_tersimpan=False):
    perlu = ruang_dibutuhkan(ukuran, sudah_tersimpan)

    def gb(n):
        return f"{n / _GIB:.1f} GB"

    persen = f"{OVERHEAD_PER_BYTE * 100:.2f}%"
    if sudah_tersimpan:
        return (f"Not enough free disk space to process this upload. The file is "
                f"already on the server. The deal index needs about {gb(perlu)} more "
                f"({persen} of the file, plus {gb(MARGIN_DISK)} for the result). "
                f"Only {gb(bebas)} is free. The upload was not deleted.")
    return (f"Not enough free disk space for this upload. The file plus its deal "
            f"index need about {gb(perlu)} free (the file, plus {persen} for the "
            f"index and {gb(MARGIN_DISK)} for the result). Only {gb(bebas)} is free. "
            f"No file was saved.")


def teks_dibersihkan(freed):
    return (f" Old data from earlier runs was already cleared before this check "
            f"({freed / _GIB:.1f} GB freed); nothing else can be removed right now.")


def cek_ruang(path, ukuran, sudah_tersimpan=False, freed=None):
    """None when the job volume can hold the run, otherwise a message for the user."""
    bebas = shutil.disk_usage(path).free
    if bebas < ruang_dibutuhkan(ukuran, sudah_tersimpan):
        pesan = pesan_disk(ukuran, bebas, sudah_tersimpan)
        return pesan + (teks_dibersihkan(freed) if freed is not None else "")
    return None


def cek_sebelum_proses(path, ukuran_sudah, ukuran_baru, freed=None):
    """Room check once staged bytes are already on this volume.

    ukuran_baru still has to be written (a small form upload). Staged bytes
    do not need to be reserved again: renaming them does not use more space.
    """
    total = max(0, ukuran_sudah) + max(0, ukuran_baru)
    perlu = ruang_dibutuhkan(total, sudah_tersimpan=True) + max(0, ukuran_baru)
    bebas = shutil.disk_usage(path).free
    if bebas < perlu:
        pesan = pesan_disk(total, bebas, sudah_tersimpan=ukuran_sudah > 0)
        return pesan + (teks_dibersihkan(freed) if freed is not None else "")
    return None


def _storage_error():
    return jsonify({"error": "The server does not have enough temporary storage for this upload. "
                             "No file was saved. Please try again after the previous upload is "
                             "finished or contact the administrator."}), 507


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


def staged_files(jobs_dir, upload_id):
    """Return the finished staged file paths in the order the user picked them.

    Returns None when the id is unknown, which the caller reports as a plain
    validation error -- a stale tab retrying an old upload must not 500.
    """
    stage = _stage_dir(jobs_dir, upload_id)
    if not stage or not stage.is_dir():
        return None
    manifest = _read_manifest(stage)
    if not manifest:
        return None
    if manifest["bytes"] != manifest.get("declared_bytes"):
        return None
    paths = []
    for entry in manifest["files"]:
        p = stage / entry["stored"]
        if not p.is_file():
            return None
        paths.append((p, entry["original"]))
    return paths or None


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
                                   "bytes": e["bytes"]} for e in manifest["files"]]})

    @app.post("/upload/begin")
    def upload_begin():
        try:
            requested = int(request.form.get("bytes", "0"))
        except ValueError:
            return jsonify({"error": "Invalid upload size."}), 400
        if requested < 0 or requested > MAX_TOTAL:
            return jsonify({"error": f"The selected files exceed the {teks_batas()} limit."}), 413
        ticket = (request.form.get("ticket") or "").strip()
        if not cleanup._valid_ticket(ticket):
            ticket = uuid.uuid4().hex
        # One worker at a time: clean, measure, then reserve. Without this two
        # workers could both see the same free space and both start.
        with cleanup.admission(jobs_dir):
            freed = cleanup.clean(jobs_dir)
            salah = cek_ruang(_stage_root(jobs_dir), requested, sudah_tersimpan=False,
                              freed=freed)
            lain = cleanup.any_activity(jobs_dir)
            berat = requested >= cleanup.HEAVY_BYTES
            sibuk = (berat and cleanup.heavy_activity(jobs_dir)) or (salah and lain)
            posisi = cleanup.queue_position(jobs_dir, ticket)
            if sibuk or (berat and posisi > 1):
                posisi = cleanup.queue_touch(jobs_dir, ticket)
                return jsonify({
                    "queued": True, "ticket": ticket, "position": posisi,
                    "message": "The server is currently processing another user's file. "
                               "Your run is queued and will start automatically when it "
                               "finishes.",
                }), 202
            cleanup.queue_leave(jobs_dir, ticket)
            if salah:
                return jsonify({"error": salah}), 507
            upload_id = uuid.uuid4().hex
            stage = _stage_root(jobs_dir) / upload_id
            try:
                sekarang = time.time()
                stage.mkdir()
                _write_manifest(stage, {"files": [], "bytes": 0,
                                        "declared_bytes": requested,
                                        "started": sekarang,
                                        "last_chunk": sekarang})
            except OSError:
                shutil.rmtree(stage, ignore_errors=True)
                return _storage_error()
        return jsonify({"id": upload_id, "freed": freed})

    @app.post("/upload/queue/leave")
    def upload_queue_leave():
        cleanup.queue_leave(jobs_dir, (request.form.get("ticket") or "").strip())
        return jsonify({"left": True})

    @app.post("/upload/discard")
    def upload_discard():
        discard(jobs_dir, request.form.get("id"))
        return jsonify({"discarded": True})

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
                return _storage_error()
            return jsonify({"received": manifest["bytes"]})
