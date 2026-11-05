#!/usr/bin/env python3
"""v1.41 resume-across-refresh: idle TTL, status hashes, staged counted in preflight,
supersede frees same-user staging, tool.html resume UX strings (no pagehide discard).

python3 test_upload_resume_v141.py
"""
import base64
import hashlib
import io
import json
import os
import tempfile
from pathlib import Path
from unittest import mock

TEST_ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TEST_ROOT / "t.db")
os.environ["DW_JOB_DIR"] = str(TEST_ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "tester"
os.environ["DW_ADMIN_PASS"] = "tester-password-1"
# Ensure defaults under test (env may be unset)
os.environ.pop("DW_UPLOAD_IDLE_SECONDS", None)

import cleanup  # noqa: E402
import chunked  # noqa: E402
from app import JOBS, app  # noqa: E402

assert cleanup.UPLOAD_IDLE_SECONDS == 4 * 3600, cleanup.UPLOAD_IDLE_SECONDS
assert chunked.STAGE_MAX_AGE_HOURS == 4, chunked.STAGE_MAX_AGE_HOURS
print("OK idle TTL", cleanup.UPLOAD_IDLE_SECONDS, "STAGE_MAX_AGE_HOURS", chunked.STAGE_MAX_AGE_HOURS)

app.config["TESTING"] = True
AUTH = "Basic " + base64.b64encode(b"tester:tester-password-1").decode()
c = app.test_client()
c.environ_base["HTTP_AUTHORIZATION"] = AUTH


def begin(total, user_client=None):
    cli = user_client or c
    r = cli.post("/upload/begin", data={"bytes": str(total)})
    assert r.status_code == 200, r.data
    return r.get_json()["id"]


def send(cli, uid, index, seq, blob, name="part.csv"):
    return cli.post("/upload/chunk", data={
        "id": uid, "name": name, "index": str(index), "seq": str(seq),
        "size": str(len(blob)), "chunk": (io.BytesIO(blob), "blob"),
    }, content_type="multipart/form-data")


# ---- status exposes hashes so the browser can validate the first chunk ----
A = b"RESUME-HEAD-" + b"A" * 2000
B = b"B" * 1500
uid = begin(len(A) + len(B))
assert send(c, uid, 0, 0, A).status_code == 200
st = c.get(f"/upload/status?id={uid}")
assert st.status_code == 200, st.data
body = st.get_json()
assert body["bytes"] == len(A)
assert body["files"][0]["chunks"] == 1
assert body["files"][0]["hashes"][0] == hashlib.sha256(A).hexdigest()[:16]
print("OK status hashes for resume", body["files"][0]["hashes"][0])

# Continue from next seq (resume after "refresh")
assert send(c, uid, 0, 1, B).status_code == 200
st2 = c.get(f"/upload/status?id={uid}").get_json()
assert st2["bytes"] == len(A) + len(B)
assert st2["files"][0]["chunks"] == 2
print("OK resume continues from next seq")

# ---- preflight counts staged bytes (same user leftover) ----
class FakeUsage:
    def __init__(self, free):
        self.free = free
        self.total = free + 10 * 1024 ** 3
        self.used = 10 * 1024 ** 3

# With staging already holding len(A)+len(B), preflight for a tiny new upload
# should report staged > 0 when disk_usage is mocked large enough.
with mock.patch("chunked.shutil.disk_usage", return_value=FakeUsage(50 * 1024 ** 3)):
    pf = c.get("/upload/preflight?bytes=1048576")
assert pf.status_code == 200, pf.data
pinfo = pf.get_json()
assert pinfo.get("staged", 0) >= len(A) + len(B), pinfo
print("OK preflight counts staged", pinfo["staged"])

# ---- new begin by same user supersedes prior staging ----
uid2 = begin(100)
assert not (JOBS / "staging" / uid).is_dir() or cleanup.superseded(JOBS / "staging" / uid) or \
    not (JOBS / "staging" / uid).exists()
# Prior id should 404 on status (gone or superseded → not usable)
st_gone = c.get(f"/upload/status?id={uid}")
assert st_gone.status_code in (404, 410), (st_gone.status_code, st_gone.data)
print("OK supersede frees prior same-user staging")


# ---- status reports ingested indexes; resume can skip them ----
from deal_segregator import _info_mt4  # noqa: F401 — ensure import works under test
import deal_segregator as ds

# Build a tiny MT5 CSV so ingest succeeds on segregate path
mt5_header = "Deal\tLogin\tName\tTime\tType\tEntry\tSymbol\tVolume\tCommission\tFee\tSwap\tProfit\tBalance\tComment\tCurrency\n"
mt5_row = "1\t100\tA\t2026.10.04 01:00:00\tbuy\tout\tEURUSD\t1.0\t0\t0\t0\t1.5\t0\t\tUSD\n"
blob_mt5 = (mt5_header + mt5_row).encode("utf-8")
uid_ing = begin(len(blob_mt5) * 2, c)
assert send(c, uid_ing, 0, 0, blob_mt5, name="part0.csv").status_code == 200
r_ing = c.post("/upload/ingest", data={"id": uid_ing, "index": "0"})
# DPM path (allow_mt4 off) — MT5 should ingest
assert r_ing.status_code == 200, r_ing.data
body_ing = r_ing.get_json()
assert body_ing.get("ingested") is True, body_ing
st_ing = c.get(f"/upload/status?id={uid_ing}").get_json()
assert 0 in st_ing.get("ingested", []), st_ing
assert st_ing["files"][0].get("ingested") is True
assert "prior.ingested" in Path("templates/tool.html").read_text() or "wasIngested" in Path("templates/tool.html").read_text()
print("OK status reports ingested; client skips re-upload of ingested parts")

# ---- tool.html: no pagehide discard; resume UX present ----
html = Path("templates/tool.html").read_text()
assert "sendBeacon('/upload/discard'" not in html and 'sendBeacon("/upload/discard"' not in html, \
    "pagehide must not discard staging in v1.41"
assert "Lanjutkan upload sebelumnya" in html
assert "Resume previous upload" in html
assert "dw_upload_resume:" in html
assert "sha256_16" in html
assert "resumeReady" in html
assert "Paused — re-select" in html or "paused until you continue" in html
assert "wasIngested" in html or "prior.ingested" in html
assert "already parsed into the index" in html
print("OK tool.html resume UX (no pagehide discard)")

# ---- VERSION ----
ver = Path("VERSION").read_text()
assert "version=1.41" in ver
assert "2026-11-05" in ver
print("OK VERSION 1.41")

print("ALL v1.41 resume checks passed")
