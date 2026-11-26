#!/usr/bin/env python3
"""DPM Deal Segregator: KYC Country + UID in front of both Monthly Summary sheets,
Total Volume last on Monthly Summary (v1.50, DPM Malaysia team via Bayu, 26 Nov 2026).

Synthetic data only: made-up logins (9001..9004), user_ids (70001..) and countries.

1. engine: match / no UID / UID without KYC row; USC ÷ 100, USD and JPY unchanged;
   FFC000 headers; every old value only shifted; Daily sheets unchanged; with FX the
   USD columns stay in place and Total Volume is after them (column R).
2. missing lookup files: still runs, both columns blank, note on Verifikasi.
3. KVB (--allow-mt4) and a plain CLI run (no --kyc-lookup) are not changed.
4. lookup module: header validation, xlsx/csv, cache by mtime, first match wins.
5. web: page status + note, upload by any DPM user (bad header rejected, good one saved
   and dated), KVB-only user refused, KVB page has no lookup card, DPM job fills the columns.
6. Login: integer with General ("no format") number format on all four DPM sheets;
   KVB keeps text; non-numeric Logins unchanged.
"""
import base64
import csv
import io
import os
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import testkit

ROOT = Path(__file__).resolve().parent
TMP = testkit.hermetic_env(admin_user="boss", admin_pass="boss-password-1")
LOOKUP = Path(os.environ["DW_DSEG_LOOKUP_DIR"])

from openpyxl import Workbook, load_workbook  # noqa: E402

import deal_segregator as ds  # noqa: E402
import dseg_lookup as L  # noqa: E402

HEADER = ["Deal", "Position", "Login", "Group", "Country", "Time", "Type", "Entry",
          "Symbol", "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]
# login -> (currency, volume per deal)
LOGINS = {9001: ("USD", "0.5"), 9002: ("USC", "50"), 9003: ("JPY", "2"), 9004: ("USD", "1")}


def write_deals(path):
    with Path(path).open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        n = 0
        for day in (1, 2):
            for login, (cur, vol) in LOGINS.items():
                for k, entry in enumerate(("in", "out")):
                    n += 1
                    w.writerow([n, f"P{n}", login, r"real\X", "", f"2026.10.{day:02d} 0{k+1}:00:00",
                                "buy", entry, "EURUSD", vol, "-1", "0", "0",
                                "10" if entry == "out" else "0", cur])
    return Path(path)


def write_xlsx(path, header, rows, title="Sheet1", extra_sheet=False):
    wb = Workbook()
    ws = wb.active
    if extra_sheet:                    # header on a later sheet must still be found
        ws.title = "notes"
        ws.append(["something else"])
        ws = wb.create_sheet(title)
    else:
        ws.title = title
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    wb.close()
    return Path(path)


def write_lookups(folder, account_ext=".xlsx"):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    acc_rows = [(70001, 9001), (70002, 9002), (70003, 9003), (70009, 5555),
                (70099, 9001)]                     # duplicate account: first one wins
    if account_ext == ".csv":
        with (folder / "account_uid.csv").open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["User_ID ", "Account"])
            w.writerows(acc_rows)
    else:
        write_xlsx(folder / "account_uid.xlsx", ["user_id", "account"], acc_rows)
    write_xlsx(folder / "uid_kyc_country.xlsx", ["user_id", "Country of KYC", "name"],
               [(70001, "zz", "x"), (70002, "yy", "y"), (70009, "ww", "z")],
               extra_sheet=True)


def books(zpath):
    out, wbs = {}, {}
    with zipfile.ZipFile(zpath) as z:
        for name in z.namelist():
            wb = load_workbook(io.BytesIO(z.read(name)))
            out[name] = {s: [list(r) for r in wb[s].iter_rows(values_only=True)]
                         for s in wb.sheetnames}
            wbs[name] = wb
    return out, wbs


def verif(book):
    return {r[0]: r[1] for r in book["Verifikasi"] if r and r[0]}


work = Path(tempfile.mkdtemp(dir=TMP))
deals = write_deals(work / "deals.csv")
MS = "Deals - Monthly Summary.xlsx"
DD = "Deals - Daily.xlsx"

# ------------------------------------------------------------ 1. engine, lookups on
lk1 = work / "lk1"
write_lookups(lk1)
ds.proses([deals], work / "base.zip")                       # v1.49 layout (no flag)
ds.proses([deals], work / "kyc.zip", kyc_lookup=str(lk1))
base, _ = books(work / "base.zip")
kyc, kwb = books(work / "kyc.zip")

out_h = kyc[MS]["Monthly Summary"][0]
in_h = kyc[MS]["Monthly Summary - In"][0]
assert out_h == ["KYC Country", "UID", "Login", "Month", "Type", "Symbol", "Deals", "Volume",
                 "Commission", "Fee", "Swap", "Profit", "Currency", "Total Volume"], out_h
assert in_h == ["KYC Country", "UID"] + base[MS]["Monthly Summary - In"][0], in_h
ws = kwb[MS]["Monthly Summary"]
fills = {c.value: c.fill.fgColor.rgb[-6:] for c in ws[1]}
assert fills["KYC Country"] == fills["UID"] == fills["Total Volume"] == "FFC000", fills
assert fills["Login"] == "1F4E78", fills
assert ws.freeze_panes == "A2"
assert kwb[MS]["Monthly Summary"]["D2"].number_format == "mmm yyyy"   # Month moved to D
assert kwb[MS]["Monthly Summary - In"]["D2"].number_format == "mmm yyyy"

# every old value is only shifted two columns to the right
for sheet in ("Monthly Summary", "Monthly Summary - In"):
    old = base[MS][sheet]
    new = kyc[MS][sheet]
    assert [r[2:2 + len(old[0])] for r in new] == old, sheet
# Daily workbooks are unchanged
for sheet in ("Daily", "Daily - In"):
    assert kyc[DD][sheet] == base[DD][sheet], sheet

rows = {r[2]: dict(zip(out_h, r)) for r in kyc[MS]["Monthly Summary"][1:]}
assert rows[9001]["UID"] == 70001 and rows[9001]["KYC Country"] == "zz"   # match; first dup wins
assert rows[9002]["UID"] == 70002 and rows[9002]["KYC Country"] == "yy"
assert rows[9003]["UID"] == 70003 and rows[9003]["KYC Country"] is None   # UID, no KYC row
assert rows[9004]["UID"] is None and rows[9004]["KYC Country"] is None    # no UID
# Total Volume: USC / 100, USD and JPY exact (2 out deals per login)
assert rows[9002]["Volume"] == 100 and rows[9002]["Total Volume"] == 1.0, rows[9002]
assert rows[9001]["Total Volume"] == rows[9001]["Volume"] == 1.0
assert rows[9003]["Total Volume"] == rows[9003]["Volume"] == 4
assert rows[9004]["Total Volume"] == rows[9004]["Volume"] == 2
ins = {r[2]: r for r in kyc[MS]["Monthly Summary - In"][1:]}
assert ins[9001][:2] == ["zz", 70001] and ins[9004][:2] == [None, None]
v = verif(kyc[MS])
assert v["Logins matched to a UID"] == "3 of 4", v
assert v["Logins with no UID (KYC Country and UID left blank)"] == "1: 9004", v
assert v["UIDs with no KYC row (KYC Country left blank)"] == "1: 70003", v
assert "updated " in v["UID / KYC lookup (KYC Country, UID columns)"]
# Login: integer + General ("no format") on every sheet, DPM only (Bayu, 26 Nov 2026)
for name, sheets in (("Daily", ("Daily", "Daily - In")),
                     ("Monthly", ("Monthly Summary", "Monthly Summary - In"))):
    wbk = kwb[DD if name == "Daily" else MS]
    for sheet in sheets:
        hdr = [c.value for c in wbk[sheet][1]]
        col = hdr.index("Login") + 1
        cells = [r[0] for r in wbk[sheet].iter_rows(min_row=2, min_col=col, max_col=col)]
        assert cells and all(isinstance(c.value, int) for c in cells), sheet
        assert all(c.number_format == "General" for c in cells), sheet
assert ds.login_sebagai_angka("503617") == 503617
for keep in ("A123", "0123", "1" * 16, "", None, 7):
    assert ds.login_sebagai_angka(keep) == keep
print("OK engine: match / no match / no KYC, USC/USD/JPY, shift-only, Daily unchanged")

# with Client Equity FX: USD columns stay, Total Volume after them (column R)
fx = work / "fx.xlsx"
write_xlsx(fx, ["date", "Currency", "rate"],
           [(f"2026-10-0{d}", c, r) for d in (1, 2) for c, r in (("USD", 1), ("JPY", 150))],
           title="Query result")
ds.proses([deals, fx], work / "fxbase.zip")
ds.proses([deals, fx], work / "fxkyc.zip", kyc_lookup=str(lk1))
fb, _ = books(work / "fxbase.zip")
fk, fwb = books(work / "fxkyc.zip")
h = fk[MS]["Monthly Summary"][0]
assert len(h) == 18 and h[17] == "Total Volume" and h[13:17] == ["Commission (USD)", "Fee (USD)",
                                                                "Swap (USD)", "Profit (USD)"], h
assert fwb[MS]["Monthly Summary"]["R1"].value == "Total Volume"
assert len(fk[MS]["Monthly Summary - In"][0]) == 17
for sheet in ("Monthly Summary", "Monthly Summary - In"):
    old = fb[MS][sheet]
    assert [r[2:2 + len(old[0])] for r in fk[MS][sheet]] == old, sheet
print("OK FX layout: 18 / 17 columns, Total Volume = R")

# interleaved (chunked) web uploads finish with --from-session: same result
import subprocess  # noqa: E402
state = work / "session"
state.mkdir()
sess = ds.IngestSession(state)
sess.ingest(deals, hapus=False, original_name="deals.csv")
sess.save()
sess.close()
r = subprocess.run([sys.executable, str(ROOT / "deal_segregator.py"), "--from-session", str(state),
                    "-o", str(work / "sess.zip"), "--kyc-lookup", str(lk1)],
                   capture_output=True, text=True, cwd=ROOT)
assert r.returncode == 0, r.stderr
assert "UID / KYC lookup (Monthly Summary): Logins matched to a UID: 3 of 4" in r.stdout, r.stdout
sb, _ = books(work / "sess.zip")
assert sb[MS]["Monthly Summary"] == kyc[MS]["Monthly Summary"]
assert sb[MS]["Monthly Summary - In"] == kyc[MS]["Monthly Summary - In"]
print("OK --from-session path")

# ------------------------------------------------------------ 2. lookup files missing
ds.proses([deals], work / "none.zip", kyc_lookup=str(work / "no-such-folder"))
nb, _ = books(work / "none.zip")
assert nb[MS]["Monthly Summary"][0][:2] == ["KYC Country", "UID"]
assert all(r[0] is None and r[1] is None for r in nb[MS]["Monthly Summary"][1:])
assert all(r[0] is None and r[1] is None for r in nb[MS]["Monthly Summary - In"][1:])
assert [r[-1] for r in nb[MS]["Monthly Summary"][1:]] == [r[-1] for r in kyc[MS]["Monthly Summary"][1:]]
note = verif(nb[MS])["UID / KYC lookup (KYC Country, UID columns)"]
assert note.startswith("NOT AVAILABLE") and "are not stored on the server" in note, note
only_acc = work / "only-acc"
write_lookups(only_acc)
(only_acc / "uid_kyc_country.xlsx").unlink()
assert L.load(only_acc)["missing"] == ["kyc"]
print("OK missing lookup files: blank columns + note")

# ------------------------------------------------------------ 3. KVB / plain CLI unchanged
ds.proses([deals], work / "kvb0.zip", allow_mt4=True)
ds.proses([deals], work / "kvb1.zip", allow_mt4=True, kyc_lookup=str(lk1))
assert books(work / "kvb0.zip")[0][MS]["Monthly Summary"] == books(work / "kvb1.zip")[0][MS]["Monthly Summary"]
assert "KYC Country" not in books(work / "kvb1.zip")[0][MS]["Monthly Summary"][0]
kvb_daily = books(work / "kvb1.zip")[0][DD]["Daily"]
assert all(isinstance(r[0], str) for r in kvb_daily[1:])          # KVB Login stays text
assert "KYC Country" not in base[MS]["Monthly Summary"][0]
print("OK KVB and a CLI run without --kyc-lookup unchanged")

# ------------------------------------------------------------ 4. lookup module
lk2 = work / "lk2"
write_lookups(lk2, account_ext=".csv")                    # csv + header case / spaces
got = L.load(lk2)
assert got["ok"] and got["data"]["account"]["map"]["9001"] == "70001"
assert got["data"]["account"]["duplicates"] == 1 and got["data"]["account"]["conflicts"] == 1
assert (lk2 / L.CACHE).is_file()
L._memo.clear()
assert L.load(lk2)["data"]["kyc"]["map"]["70001"] == "zz"           # from the pickle
time.sleep(0.01)
write_xlsx(lk2 / "uid_kyc_country.xlsx", ["user_id", "Country of KYC"], [(70001, "qq")])
assert L.load(lk2)["data"]["kyc"]["map"] == {"70001": "qq"}         # mtime changed -> rebuilt
assert L.norm_id(9001) == L.norm_id("9001") == L.norm_id(9001.0) == L.norm_id(" 9001.0 ") == "9001"
bad = write_xlsx(work / "bad.xlsx", ["user_id", "acct"], [(1, 2)])
for kind, path in (("account", bad), ("kyc", bad)):
    try:
        L.save(kind, path, "bad.xlsx", folder=lk2)
        raise AssertionError("bad header accepted")
    except ValueError as exc:
        assert "header must have the columns" in str(exc), exc
assert L.load(lk2)["data"]["kyc"]["map"] == {"70001": "qq"}         # stored file kept
try:
    L.save("account", bad, "bad.txt", folder=lk2)
    raise AssertionError("bad extension accepted")
except ValueError as exc:
    assert ".xlsx, .xlsm or .csv" in str(exc)
L.save("account", work / "lk1" / "account_uid.xlsx", "x.xlsx", folder=lk2)  # replaces csv
assert not (lk2 / "account_uid.csv").exists() and (lk2 / "account_uid.xlsx").exists()
print("OK lookup module: validation, csv/xlsx, cache, first match")

# ------------------------------------------------------------ 5. web
import app as A  # noqa: E402

testkit.assert_isolated(TMP)
assert A.dseg_lookup.lookup_dir() == LOOKUP
A.app.config["TESTING"] = True
c = A.app.test_client()
c.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(b"boss:boss-password-1").decode()
c.environ_base["HTTP_ACCEPT"] = "text/html"

page = c.get("/tool/segregate").data.decode()
assert "UID / KYC lookup files" in page and 'id="kyc-lookup-missing"' in page
assert "Update UID / KYC lookup files" in page and 'action="/tool/segregate/lookup"' in page
assert "not on the server" in page
assert "UID / KYC lookup files" not in c.get("/kvb/tool/segregate").data.decode()
assert "UID / KYC lookup files" not in c.get("/tool/segregated").data.decode()

def up(**files):
    data = {k: (io.BytesIO(Path(p).read_bytes()), Path(p).name) for k, p in files.items()}
    return c.post("/tool/segregate/lookup", data=data, content_type="multipart/form-data")

r = up(account_file=bad)
assert r.status_code == 400 and b"header must have the columns user_id, account" in r.data, r.data[-800:]
assert not any(LOOKUP.glob("account_uid.*"))
r = c.post("/tool/segregate/lookup", data={}, content_type="multipart/form-data")
assert r.status_code == 400 and b"Choose the account file" in r.data
r = up(account_file=lk1 / "account_uid.xlsx", kyc_file=lk1 / "uid_kyc_country.xlsx")
assert r.status_code == 200 and b"Lookup files updated" in r.data, r.data[-800:]
html = r.data.decode()
assert 'id="kyc-lookup-missing"' not in html
assert re.search(r"account_uid\.xlsx</strong>, updated \d{4}-\d{2}-\d{2} \d{2}:\d{2}", html)
assert (LOOKUP / "account_uid.xlsx").is_file() and (LOOKUP / "uid_kyc_country.xlsx").is_file()

# every DPM user (not only admins) may replace the files (Bayu, 26 Nov 2026)
def user_client(name, **access):
    r = c.post("/admin/users", data={"username": name, "password": "ops-password-1", **access})
    assert r.status_code in (302, 303), r.status_code
    cl = A.app.test_client()
    cl.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(
        f"{name}:ops-password-1".encode()).decode()
    cl.environ_base["HTTP_ACCEPT"] = "text/html"
    return cl

c2 = user_client("ops1", can_dpm="1")
p2 = c2.get("/tool/segregate").data.decode()
assert "UID / KYC lookup files" in p2 and 'action="/tool/segregate/lookup"' in p2
r = c2.post("/tool/segregate/lookup", data={"kyc_file": (io.BytesIO(b"user_id,acct\n1,aa\n"), "k.csv")},
            content_type="multipart/form-data")
assert r.status_code == 400 and b"user_id, Country of KYC" in r.data      # header still checked
assert L.load(LOOKUP)["data"]["kyc"]["map"]["70001"] == "zz"
kyc_csv = b"user_id,Country of KYC\n70001,zz\n70002,yy\n70009,ww\n"
r = c2.post("/tool/segregate/lookup", data={"kyc_file": (io.BytesIO(kyc_csv), "k.csv")},
            content_type="multipart/form-data")
assert r.status_code == 200 and b"Lookup files updated" in r.data, r.data[-600:]
assert (LOOKUP / "uid_kyc_country.csv").is_file() and not (LOOKUP / "uid_kyc_country.xlsx").exists()
assert L.load(LOOKUP)["data"]["kyc"]["map"]["70002"] == "yy"

# a KVB-only user has no DPM access: no page, no upload
c3 = user_client("kvbops", can_kvb="1")
assert c3.get("/tool/segregate").status_code in (302, 303, 401, 403)
r = c3.post("/tool/segregate/lookup", data={"kyc_file": (io.BytesIO(kyc_csv), "x.csv")},
            content_type="multipart/form-data")
assert r.status_code in (302, 303, 401, 403), r.status_code
assert (LOOKUP / "uid_kyc_country.csv").read_bytes() == kyc_csv

# DPM web job: the runner passes --kyc-lookup with the server folder
r = c.post("/tool/segregate", data={"file": (deals.open("rb"), "deals.csv")},
           content_type="multipart/form-data")
assert r.status_code == 303, (r.status_code, r.data[:300])
job_id = r.headers["Location"].rstrip("/").rsplit("/", 1)[-1]
for _ in range(600):
    info = A._read_state(job_id)
    if info and info["state"] in ("done", "failed"):
        break
    time.sleep(0.2)
assert info["state"] == "done", info
dl = c.get(f"/job/{job_id}/download")
wz = work / "web.zip"
wz.write_bytes(dl.data)
wb_, _ = books(wz)
web_rows = {r[2]: r for r in wb_[MS]["Monthly Summary"][1:]}
assert web_rows[9001][:2] == ["zz", 70001] and web_rows[9002][-1] == 1.0
assert wb_[DD]["Daily"] == base[DD]["Daily"]
assert "UID / KYC lookup (Monthly Summary): Logins matched to a UID: 3 of 4" in (info.get("log") or ""), info.get("log")
src = (ROOT / "app.py").read_text(encoding="utf-8")
assert src.count('["--kyc-lookup", str(dseg_lookup.lookup_dir())]') == 2
print("OK web: status, upload by any DPM user, KVB-only user refused, KVB page unchanged, DPM job filled")

# ------------------------------------------------------------ version
ver = (ROOT / "VERSION").read_text(encoding="utf-8")
assert "version=1.50" in ver and "build=2026-11-26" in ver, ver
assert "v1.50" in (ROOT / "README.md").read_text(encoding="utf-8")
assert "dseg_lookup/" in (ROOT / ".gitignore").read_text(encoding="utf-8")
print("test_deal_seg_kyc_v150: ok")
