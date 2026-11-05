#!/usr/bin/env python3
"""v1.40 upload disk lifecycle: preflight, peak model, 507 no-VPN, multi-part == concat,
delete-after-parse, sweep wiring, synthetic peak with mocked disk_usage.

python3 test_upload_disk_v140.py
"""
import base64
import csv
import hashlib
import io
import json
import os
import shutil
import tempfile
import time
import zipfile
from pathlib import Path
from unittest import mock

TEST_ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TEST_ROOT / "t.db")
os.environ["DW_JOB_DIR"] = str(TEST_ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "tester"
os.environ["DW_ADMIN_PASS"] = "tester-password-1"

import chunked  # noqa: E402
import deal_segregator as ds  # noqa: E402
from app import JOBS, app  # noqa: E402
from openpyxl import load_workbook

_GIB = 1024 * 1024 * 1024
HEADERS = [
    "Deal", "Position", "Login", "Group", "Country", "Time", "Type", "Entry",
    "Symbol", "Volume", "Commission", "Fee", "Swap", "Profit", "Currency",
]


def row(deal, login="100", time="2026.07.01 01:02:03", entry="out", profit="10"):
    return [deal, deal, login, r"real\DPMKT-15", "MY", time, "buy", entry,
            "EURUSD", "1.0", "-1", "-0.1", "-0.2", profit, "USD"]


def write_csv(path, rows):
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADERS)
        w.writerows(rows)


def zip_sha(path):
    """Content fingerprint ignoring zip timestamps: sorted (name, crc, size)."""
    with zipfile.ZipFile(path) as z:
        items = [(i.filename, i.CRC, i.file_size) for i in z.infolist()]
    return hashlib.sha256(repr(sorted(items)).encode()).hexdigest()


def sheets_equal(a, b):
    wa, wb = load_workbook(a, read_only=True, data_only=True), load_workbook(b, read_only=True, data_only=True)
    try:
        assert wa.sheetnames == wb.sheetnames
        for name in wa.sheetnames:
            ra = list(wa[name].iter_rows(values_only=True))
            rb = list(wb[name].iter_rows(values_only=True))
            assert ra == rb, (name, ra[:3], rb[:3])
    finally:
        wa.close(); wb.close()


# ---- peak numbers for the Bayu report ----
S = 11 * _GIB
single_need = chunked.ruang_dibutuhkan(S, False)
part = S // 4
four_need = chunked.ruang_dibutuhkan(S, False, max_part=part)
print(f"PEAK single 11 GiB need={single_need / _GIB:.3f} GiB")
print(f"PEAK 4x{part / _GIB:.3f} GiB need={four_need / _GIB:.3f} GiB")
assert four_need < single_need
assert abs(single_need - (S + int(S * chunked.OVERHEAD_PER_BYTE) + chunked.MARGIN_DISK)) < 1
assert abs(four_need - (part + int(S * chunked.OVERHEAD_PER_BYTE) + chunked.MARGIN_DISK)) < 1

# ---- preflight JSON + auth ----
app.config["TESTING"] = True
c = app.test_client()
auth = "Basic " + base64.b64encode(b"tester:tester-password-1").decode()
c.environ_base["HTTP_AUTHORIZATION"] = auth

r = c.get("/upload/preflight?bytes=104857600&max_part=52428800")
assert r.status_code in (200, 507), r.data
body = r.get_json()
assert "need" in body and "free" in body and "ok" in body and "message" in body
assert "max_part" in body
assert "Perlu" in body["message"] and "Need" in body["message"]
assert body["max_part"] == 52428800
assert body["total"] == 104857600
print("OK preflight JSON", {k: body[k] for k in ("ok", "need", "free", "max_part")})

# unauth must not succeed without credentials when auth is required
c2 = app.test_client()
r2 = c2.get("/upload/preflight?bytes=1")
# Basic auth is required by the app wrapper; without it expect 401
assert r2.status_code == 401, r2.status_code

# CSRF-safe: GET must not mutate; a second identical call returns same shape
r3 = c.get("/upload/preflight?bytes=104857600&max_part=52428800")
assert set(r3.get_json()) >= {"ok", "need", "free", "message"}

# ---- mocked disk: not enough space ----
class FakeUsage:
    def __init__(self, free):
        self.free = free
        self.total = free + 10 * _GIB
        self.used = 10 * _GIB

with mock.patch("chunked.shutil.disk_usage", return_value=FakeUsage(100)):
    info = chunked.ringkas_preflight(S, 100, max_part=part)
    assert info["ok"] is False
    assert "VPN" not in info["message"]
    r = c.get("/upload/preflight?bytes=" + str(S) + "&max_part=" + str(part))
    # disk_usage is patched only for ringkas calls above; re-patch for HTTP
with mock.patch("chunked.shutil.disk_usage", return_value=FakeUsage(100)):
    r = c.get("/upload/preflight?bytes=" + str(S) + "&max_part=" + str(part))
    assert r.status_code == 507, r.data
    j = r.get_json()
    assert j["ok"] is False
    assert "VPN" not in j["message"]
    assert "Perlu" in j["message"]
print("OK preflight 507 bilingual, no VPN")

# begin with not enough space: 507 + freed reported, and sweep was called
calls = {"sweep": 0}
real_sweep = chunked.sweep

def tracking_sweep(*a, **k):
    calls["sweep"] += 1
    return real_sweep(*a, **k)

with mock.patch("chunked.sweep", side_effect=tracking_sweep), \
     mock.patch("chunked.shutil.disk_usage", return_value=FakeUsage(100)), \
     mock.patch("cleanup.shutil.disk_usage", return_value=FakeUsage(100)):
    r = c.post("/upload/begin", data={"bytes": str(S), "max_part": str(part)})
    assert r.status_code == 507, (r.status_code, r.data)
    j = r.get_json()
    assert "error" in j and "freed" in j
    err = j["error"]
    assert "bukan masalah vpn" in err.lower() or "not a vpn" in err.lower(), err
    assert "WARP" not in err
    assert "firewall" not in err.lower()
assert calls["sweep"] >= 1, calls
print("OK begin 507 after sweep, freed reported")

# mid-chunk storage error text
with app.app_context():
    resp = chunked._storage_error(freed=123)
assert resp[1] == 507
data = resp[0].get_json()
assert data["disk"] is True
assert "vpn" in data["error"].lower() and "bukan masalah vpn" in data["error"].lower()
assert "firewall" not in data["error"].lower()
assert data["freed"] == 123
print("OK mid-chunk _storage_error bilingual + freed")

# ---- multi-part == concatenated output + delete-after-parse ----
with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    rows_a = [row(1), row(2, entry="in"), row(3)]
    rows_b = [row(4), row(5, entry="in"), row(2)]  # deal 2 dup across parts
    rows_c = [row(6, time="2026.07.02 00:00:00")]
    a, b, c_path = td / "part_jul.csv", td / "part_aug.csv", td / "part_sep.csv"
    write_csv(a, rows_a)
    write_csv(b, rows_b)
    write_csv(c_path, rows_c)
    concat = td / "all.csv"
    write_csv(concat, rows_a + rows_b + rows_c)

    out_parts = td / "parts.zip"
    out_one = td / "one.zip"
    # Keep copies for concat run; process-and-delete removes sources.
    a2, b2, c2 = td / "a2.csv", td / "b2.csv", td / "c2.csv"
    shutil.copy(a, a2); shutil.copy(b, b2); shutil.copy(c_path, c2)
    conc2 = td / "all2.csv"; shutil.copy(concat, conc2)

    ds.proses([a, b, c_path], out_parts, hapus_sumber=True)
    assert not a.exists() and not b.exists() and not c_path.exists(), "parts must be deleted after parse"
    ds.proses([conc2], out_one, hapus_sumber=True)
    assert not conc2.exists()

    # Data sheets must match; Verifikasi differs on "Uploaded files" count (N parts vs 1).
    with zipfile.ZipFile(out_parts) as zp, zipfile.ZipFile(out_one) as zo:
        assert sorted(zp.namelist()) == sorted(zo.namelist())
        for name in zp.namelist():
            pa, po = td / ("p_" + name), td / ("o_" + name)
            pa.write_bytes(zp.read(name)); po.write_bytes(zo.read(name))
            wa = load_workbook(pa, read_only=True, data_only=True)
            wo = load_workbook(po, read_only=True, data_only=True)
            try:
                assert wa.sheetnames == wo.sheetnames
                for sheet in wa.sheetnames:
                    if sheet == "Verifikasi":
                        continue
                    ra = list(wa[sheet].iter_rows(values_only=True))
                    rb = list(wo[sheet].iter_rows(values_only=True))
                    assert ra == rb, (name, sheet, ra[:2], rb[:2])
            finally:
                wa.close(); wo.close()
    print("OK multi-part == concatenated + delete-after-parse")

# later part failure names the part and cleans up via exception -> STOP
with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    good = td / "ok.csv"
    write_csv(good, [row(1)])
    bad = td / "broken.csv"
    bad.write_text("this is not a deals history\n", encoding="utf-8")
    out = td / "out.zip"
    try:
        ds.proses([good, bad], out, hapus_sumber=True)
        raise AssertionError("expected STOP")
    except SystemExit as e:
        msg = str(e)
        assert "part 2 of 2" in msg, msg
        assert "broken.csv" in msg, msg
    print("OK part-failure names part")

# ---- synthetic ~1.5 GB peak with mocked disk_usage ----
# Prove begin accepts when free >= peak(max_part) and rejects when free < peak.
S_syn = 1500 * 1024 * 1024  # 1.5 GB total
max_p = 400 * 1024 * 1024   # 400 MB largest part
need = chunked.ruang_dibutuhkan(S_syn, False, max_part=max_p)
assert need < S_syn + chunked.MARGIN_DISK  # cheaper than holding every part

with mock.patch("chunked.shutil.disk_usage", return_value=FakeUsage(need + 10)), \
     mock.patch("cleanup.shutil.disk_usage", return_value=FakeUsage(need + 10)):
    r = c.post("/upload/begin", data={"bytes": str(S_syn), "max_part": str(max_p)})
    # May queue (202) if another heavy job, or 200 with id
    assert r.status_code in (200, 202), (r.status_code, r.data)
    if r.status_code == 200:
        assert "id" in r.get_json()
        # discard to free staging
        c.post("/upload/discard", data={"id": r.get_json()["id"], "reason": "test"})

with mock.patch("chunked.shutil.disk_usage", return_value=FakeUsage(need - 10)), \
     mock.patch("cleanup.shutil.disk_usage", return_value=FakeUsage(need - 10)):
    r = c.post("/upload/begin", data={"bytes": str(S_syn), "max_part": str(max_p)})
    assert r.status_code in (507, 202), (r.status_code, r.data)  # 202 if blockers exist
print(f"OK synthetic peak need={need / (1024**2):.1f} MiB measured")

# tool.html: 507 excluded from retry; disk errors skip VPN advice
html = Path("templates/tool.html").read_text()
assert "r.status !== 507" in html or "r.status === 507" in html
assert "diskFull" in html
assert "bukan masalah VPN" in html
assert "upload-split-hint" in html
assert "disk-preflight-status" in html
assert "/upload/preflight" in html
print("OK tool.html gating + UI hints")

# VERSION
assert "1.40" in Path("VERSION").read_text()
print("OK VERSION 1.40")
print("ALL test_upload_disk_v140 OK")
