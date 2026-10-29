"""Per-run file limit: 65 files accepted, 71 rejected; every upload field counts."""
import io
import subprocess
import sys
import tempfile
from pathlib import Path

import testkit as T

# Assign (not setdefault): an exported server DW_* env must never leak in.
_ROOT = T.hermetic_env("tester", "tester-password-1")
import app as web  # noqa: E402
T.assert_isolated(_ROOT)
import chunked  # noqa: E402
import deal_segregator  # noqa: E402
from upload_limits import MAX_FILES  # noqa: E402

assert MAX_FILES == 70
assert chunked.MAX_FILES == deal_segregator.MAX_FILES == MAX_FILES
assert chunked.teks_batas() == "12 GB", chunked.teks_batas()  # total limit unchanged

c = web.app.test_client()
c.environ_base["HTTP_AUTHORIZATION"] = "Basic " + __import__("base64").b64encode(
    b"tester:tester-password-1").decode()
page = c.get("/tool/segregate")
assert page.status_code == 200, page.status_code
if True:
    assert b"70 files maximum" in page.data and b"var MAX_FILES = 70;" in page.data
    assert b"70-file maximum" in page.data and b"50-file" not in page.data


def post(n, accounts=False, url="/tool/segregate"):
    data = {"file": [(io.BytesIO(b""), f"{i}.csv") for i in range(n)]}
    if accounts:
        data["mt4_accounts"] = (io.BytesIO(b"account,account_type\n"), "types.csv")
    return c.post(url, data=data, content_type="multipart/form-data")


assert page.status_code == 200, page.status_code
if True:
    # 0-byte files: passing the count check lands on the empty-file check, never the run.
    ok = post(65)
    assert b"no more than 70 files" not in ok.data and b"Empty files cannot be processed:" in ok.data, ok.data[:300]
    ok = post(70)
    assert b"no more than 70 files" not in ok.data and b"Empty files cannot be processed:" in ok.data
    bad = post(71)
    assert bad.status_code == 400 and b"no more than 70 files" in bad.data

kvb = c.get("/kvb/tool/segregate")
assert kvb.status_code == 200, kvb.status_code
if True:
    assert b"var MAX_FILES = 70;" in kvb.data and b"mt4-accounts" in kvb.data
    r = post(69, True, "/kvb/tool/segregate")
    assert b"no more than 70 files" not in r.data and b"Empty files cannot be processed:" in r.data
    bad = post(70, True, "/kvb/tool/segregate")    # 70 Deals/FX + account types = 71
    assert bad.status_code == 400 and b"no more than 70 files" in bad.data

# chunk index guard
begin = c.post("/upload/begin", data={"total": "10", "count": "1"})
assert begin.status_code == 200, begin.data
if True:
    uid = begin.get_json()["id"]
    r = c.post("/upload/chunk", data={"id": uid, "name": "a.csv", "index": "70", "seq": "0",
                                      "chunk": (io.BytesIO(b"x"), "blob")},
               content_type="multipart/form-data")
    assert r.status_code == 400 and b"At most 70 files" in r.data, r.data

# CLI
with tempfile.TemporaryDirectory() as d:
    paths = []
    for i in range(71):
        p = Path(d) / f"{i}.csv"; p.write_text("x"); paths.append(str(p))
    r = subprocess.run([sys.executable, "deal_segregator.py", *paths, "-o", f"{d}/o.zip"],
                       capture_output=True, text=True, cwd=Path(__file__).parent)
    assert r.returncode != 0 and "no more than 70 files" in r.stderr, r.stderr[-300:]
    r = subprocess.run([sys.executable, "deal_segregator.py", *paths[:65], "-o", f"{d}/o.zip"],
                       capture_output=True, text=True, cwd=Path(__file__).parent)
    assert "no more than" not in r.stderr, r.stderr[-300:]
print("test_upload_limit OK")
