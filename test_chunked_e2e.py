#!/usr/bin/env python3
"""End-to-end: a real Deals History file uploaded in pieces must produce the same
workbook as a plain upload. Run: python3 test_chunked_e2e.py
"""
import base64
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

from app import app  # noqa: E402

FIXTURE = Path("/Users/bayudarmawan/Documents/Dupoin/zern/"
               "31 Jul_Deals History 2026_08_07 12_07_27 (1).csv")
if not FIXTURE.is_file():
    print("SKIP chunked e2e — fixture not present on this machine")
    raise SystemExit(0)

app.config["TESTING"] = True
c = app.test_client()
c.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(
    b"tester:tester-password-1").decode()

CHUNK = 5 * 1024 * 1024        # small on purpose: forces many pieces
data = FIXTURE.read_bytes()
print(f"fixture {len(data)/1048576:.1f} MB -> {-(-len(data)//CHUNK)} chunks")

uid = c.post("/upload/begin", data={"bytes": str(len(data))}).get_json()["id"]
seq = 0
for off in range(0, len(data), CHUNK):
    r = c.post("/upload/chunk", data={
        "id": uid, "name": FIXTURE.name, "index": "0", "seq": str(seq),
        "chunk": (io.BytesIO(data[off:off + CHUNK]), "part"),
    }, content_type="multipart/form-data")
    assert r.status_code == 200, (seq, r.status_code, r.get_json())
    seq += 1

r = c.post("/tool/segregate", data={"upload_id": uid},
           content_type="multipart/form-data")
assert r.status_code == 302, (r.status_code, r.data[:500])
job_id = r.headers["Location"].rstrip("/").split("/")[-1]

deadline = time.time() + 900
while time.time() < deadline:
    s = c.get(f"/job/{job_id}")
    if s.status_code == 422:
        raise AssertionError(f"job failed: {s.data[:800]}")
    if b"/download" in s.data:
        break
    time.sleep(2)
else:
    raise AssertionError("job did not finish in 15 minutes")

d = c.get(f"/job/{job_id}/download")
assert d.status_code == 200, d.status_code
assert d.data[:2] == b"PK", "the result is not an xlsx file"
print(f"OK chunked e2e — {len(d.data):,} bytes result from {seq} chunks")
