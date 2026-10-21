#!/usr/bin/env python3
"""Upload resilience (21 Oct 2026 incident): idempotent chunk re-sends,
/upload/status for resume, incomplete bodies, and the /upload/* request log.

python3 test_upload_resilience.py
"""
import base64
import io
import json
import logging
import os
import tempfile
from pathlib import Path

TEST_ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TEST_ROOT / "t.db")
os.environ["DW_JOB_DIR"] = str(TEST_ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "tester"
os.environ["DW_ADMIN_PASS"] = "tester-password-1"

import chunked  # noqa: E402
from app import JOBS, app  # noqa: E402

app.config["TESTING"] = True
AUTH = "Basic " + base64.b64encode(b"tester:tester-password-1").decode()
c = app.test_client()
c.environ_base["HTTP_AUTHORIZATION"] = AUTH


class Capture(logging.Handler):
    def __init__(self):
        super().__init__()
        self.lines = []

    def emit(self, record):
        self.lines.append(record.getMessage())


cap = Capture()
chunked.log.addHandler(cap)


def begin(total):
    r = c.post("/upload/begin", data={"bytes": str(total)})
    assert r.status_code == 200, r.data
    return r.get_json()["id"]


def send(uid, index, seq, blob, name="MT4_RawReport_20260916.csv", size=True, headers=None):
    data = {"id": uid, "name": name, "index": str(index), "seq": str(seq),
            "chunk": (io.BytesIO(blob), "blob")}
    if size:
        data["size"] = str(len(blob))
    return c.post("/upload/chunk", data=data, content_type="multipart/form-data",
                  headers=headers or {})


def manifest(uid):
    return json.loads((JOBS / "staging" / uid / "manifest.json").read_text())


def stored(uid, index=0):
    return (JOBS / "staging" / uid / f"part-{index:02d}.csv").read_bytes()


A, B, C2 = b"A" * 1000 + b"SECRET-ROW", b"B" * 1000, b"C" * 500
uid = begin(len(A) + len(B) + len(C2))

# --- idempotent duplicate chunks ---------------------------------------------------
r = send(uid, 0, 0, A)
assert r.status_code == 200 and r.get_json() == {"received": len(A)}, r.data
r = send(uid, 0, 0, A)                 # the ack was lost: same chunk again
assert r.status_code == 200 and r.get_json() == {"received": len(A), "duplicate": True}, r.data
assert stored(uid) == A, "a repeated chunk must not be written twice"
assert send(uid, 0, 1, B).status_code == 200
assert send(uid, 0, 0, A).get_json() == {"received": len(A) + len(B), "duplicate": True}
assert send(uid, 0, 1, B).get_json()["duplicate"] is True
assert stored(uid) == A + B
# same seq, different bytes: not a duplicate
r = send(uid, 0, 0, b"Z" * len(A))
assert r.status_code == 409 and r.get_json()["expected"] == 2
# a gap is still out of order
r = send(uid, 0, 5, C2)
assert r.status_code == 409 and r.get_json()["expected"] == 2
m = manifest(uid)
assert m["bytes"] == len(A) + len(B)
assert m["files"][0]["chunks"] == 2 and len(m["files"][0]["hashes"]) == 2
assert m["files"][0]["sizes"] == [len(A), len(B)]
print("OK idempotent duplicate chunks")

# a manifest written before hashes were kept: the last chunk can still be matched
old = begin(len(A) + len(B))
assert send(old, 0, 0, A).status_code == 200
mm = manifest(old)
for e in mm["files"]:
    e.pop("hashes"); e.pop("sizes")
(JOBS / "staging" / old / "manifest.json").write_text(json.dumps(mm))
assert send(old, 0, 0, A).get_json().get("duplicate") is True
assert send(old, 0, 1, B).status_code == 200
assert send(old, 0, 0, A).status_code == 409    # older chunk, no hashes: cannot prove it
print("OK legacy manifest")

# --- /upload/status ------------------------------------------------------------
r = c.get(f"/upload/status?id={uid}")
assert r.status_code == 200, r.data
st = r.get_json()
assert st["bytes"] == len(A) + len(B) and st["declared_bytes"] == len(A) + len(B) + len(C2)
assert st["files"] == [{"index": 0, "chunks": 2, "bytes": len(A) + len(B)}], st
assert "hashes" not in json.dumps(st) and "original" not in st["files"][0]
assert c.get("/upload/status?id=" + "0" * 32).status_code == 404
assert c.get("/upload/status?id=../../etc").status_code == 404
assert app.test_client().get(f"/upload/status?id={uid}").status_code == 401, "needs login"
print("OK /upload/status")

# --- incomplete bodies -------------------------------------------------------------
r = send(uid, 0, 2, C2[:100], size=False)
assert r.status_code == 200          # no size field: older pages still work
uid2 = begin(len(C2))
r = c.post("/upload/chunk", data={"id": uid2, "name": "x.csv", "index": "0", "seq": "0",
                                  "size": str(len(C2) + 10), "chunk": (io.BytesIO(C2), "blob")},
           content_type="multipart/form-data")
assert r.status_code == 408 and r.get_json()["retry"] is True, r.data
assert manifest(uid2)["bytes"] == 0, "a short chunk must not be stored"
# body cut off mid-way: Content-Length promised more than arrived
full = (b"--B\r\nContent-Disposition: form-data; name=\"id\"\r\n\r\n" + uid2.encode() +
        b"\r\n--B\r\nContent-Disposition: form-data; name=\"chunk\"; filename=\"blob\"\r\n"
        b"Content-Type: application/octet-stream\r\n\r\n" + C2 + b"\r\n--B--\r\n")
r = c.post("/upload/chunk", input_stream=io.BytesIO(full[:60]),
           content_type="multipart/form-data; boundary=B",
           headers={"Content-Length": str(len(full))})
assert r.status_code == 408 and r.get_json()["retry"] is True, (r.status_code, r.data)
assert "expired" not in r.get_json()["error"]
print("OK incomplete bodies")

# --- request log ----------------------------------------------------------------
cap.lines.clear()
u3 = begin(len(A))
r = send(u3, 0, 0, A, headers={"CF-Connecting-IP": "203.0.113.9"})
send(u3, 0, 0, A, headers={"CF-Connecting-IP": "203.0.113.9"})
c.get(f"/upload/status?id={u3}")
app.test_client().post("/upload/chunk", data={"id": u3})       # not signed in
lines = [ln for ln in cap.lines if ln.startswith("upload ")]
assert len(lines) == 5, lines
b, ch, dup, stt, anon = lines
assert " POST /upload/begin 200 " in b
for part in (" POST /upload/chunk 200 ", f"id={u3}", "file=0", "seq=0", "ip=203.0.113.9",
             "user=tester", "bytes=", "ms="):
    assert part in ch, (part, ch)
assert "note=duplicate chunk, already stored" in dup, dup
assert " GET /upload/status 200 " in stt and f"id={u3}" in stt
assert " POST /upload/chunk 401 " in anon and "user=-" in anon and "id=-" in anon, anon
assert all("SECRET-ROW" not in ln and "AAAA" not in ln for ln in cap.lines), "file contents logged"
import re  # noqa: E402
assert re.match(r"upload \d{4}-\d\d-\d\dT\d\d:\d\d:\d\d[+-]\d\d:\d\d ", ch), ch
cap.lines.clear()
c.post("/upload/chunk", input_stream=io.BytesIO(full[:60]),
       content_type="multipart/form-data; boundary=B", headers={"Content-Length": str(len(full))})
assert any(" 408 " in ln and "note=incomplete body: Content-Length" in ln for ln in cap.lines), cap.lines
cap.lines.clear()
c.post("/upload/chunk", data={"id": uid2, "name": "x.csv", "index": "0", "seq": "0",
                              "size": "999999", "chunk": (io.BytesIO(C2), "blob")},
       content_type="multipart/form-data")
assert any("note=short read: 500 of 999999 bytes" in ln for ln in cap.lines), cap.lines
# the handler writes to stderr (pm2 error log), and only /upload/* is logged
assert any(isinstance(h, logging.StreamHandler) and getattr(h, "stream", None) is __import__("sys").stderr
           for h in chunked.log.handlers)
cap.lines.clear()
c.get("/version")
assert not cap.lines
print("OK request log")

# --- the page ships the new uploader ----------------------------------------------
html = c.get("/kvb/tool/segregate", headers={"Accept": "text/html"}).data.decode()
for needle in ("makeChunkSender", "redirect: 'manual'", "/upload/status?id=",
               "fd.append('size', piece.size)", "Connection lost", "Session expired"):
    assert needle in html, needle
print("OK upload resilience")
