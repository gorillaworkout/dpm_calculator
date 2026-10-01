#!/usr/bin/env python3
"""Chunked upload staging.

Cloudflare rejects any single request body over 100 MB, so a 900 MB batch can
never arrive in one POST. The browser therefore slices each file into pieces and
sends them one at a time; every request stays far below the edge limit while the
total is unlimited. The pieces are appended to a staging file on disk, so the
server never holds a whole upload in memory.

Flow:
    POST /upload/begin                  -> {"id": "<hex>"}
    POST /upload/chunk  (repeat)        -> {"received": <bytes>}
    POST /tool/<slug>   upload_id=<hex> -> the normal pipeline runs
"""
import fcntl
import json
import shutil
import time
import uuid
from pathlib import Path

from flask import jsonify, request
from werkzeug.utils import secure_filename

CHUNK_SUFFIXES = (".csv", ".xlsx", ".xlsm")
MAX_TOTAL = 5 * 1024 * 1024 * 1024
MAX_FILES = 50
STAGE_MAX_AGE_HOURS = 1
FREE_SPACE_RESERVE = 1024 * 1024 * 1024


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


def sweep(jobs_dir, max_age_hours=STAGE_MAX_AGE_HOURS):
    """Delete abandoned staging directories (browser closed mid-upload)."""
    cutoff = time.time() - max_age_hours * 3600
    root = _stage_root(jobs_dir)
    for entry in root.iterdir():
        try:
            manifest = _manifest_path(entry)
            activity = max(entry.stat().st_mtime,
                           manifest.stat().st_mtime if manifest.exists() else 0)
            if activity < cutoff:
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


def register(app, jobs_dir):
    """Attach the two upload endpoints. Auth is handled by the app's global hook."""

    @app.post("/upload/begin")
    def upload_begin():
        sweep(jobs_dir)
        try:
            requested = int(request.form.get("bytes", "0"))
        except ValueError:
            return jsonify({"error": "Invalid upload size."}), 400
        if requested < 0 or requested > MAX_TOTAL:
            return jsonify({"error": "The selected files exceed the 5 GB limit."}), 413
        if shutil.disk_usage(_stage_root(jobs_dir)).free < requested + FREE_SPACE_RESERVE:
            return _storage_error()
        upload_id = uuid.uuid4().hex
        stage = _stage_root(jobs_dir) / upload_id
        try:
            stage.mkdir()
            _write_manifest(stage, {"files": [], "bytes": 0, "started": time.time()})
        except OSError:
            shutil.rmtree(stage, ignore_errors=True)
            return _storage_error()
        return jsonify({"id": upload_id})

    @app.post("/upload/discard")
    def upload_discard():
        discard(jobs_dir, request.form.get("id"))
        return jsonify({"discarded": True})

    @app.post("/upload/chunk")
    def upload_chunk():
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

        blob = request.files.get("chunk")
        if not blob:
            return jsonify({"error": "Malformed upload request."}), 400
        data = blob.read()

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
            if seq == expected - 1 and target.is_file() and target.stat().st_size >= len(data):
                with open(target, "rb") as fh:
                    fh.seek(-len(data), 2)
                    if fh.read() == data:
                        return jsonify({"received": manifest["bytes"]})
            if seq != expected:
                return jsonify({"error": "Chunk arrived out of order.", "expected": expected}), 409
            if manifest["bytes"] + len(data) > MAX_TOTAL:
                shutil.rmtree(stage, ignore_errors=True)
                return jsonify({"error": "The files add up to more than 5 GB. "
                                         "Please split them into several smaller runs."}), 413

            try:
                with open(target, "ab") as fh:
                    fh.write(data)
                if entry:
                    entry["chunks"] += 1
                    entry["bytes"] += len(data)
                else:
                    manifest["files"].append({"index": index, "original": original,
                                              "stored": stored, "chunks": 1, "bytes": len(data)})
                manifest["files"].sort(key=lambda e: e["index"])
                manifest["bytes"] += len(data)
                _write_manifest(stage, manifest)
            except OSError:
                shutil.rmtree(stage, ignore_errors=True)
                return _storage_error()
            return jsonify({"received": manifest["bytes"]})
