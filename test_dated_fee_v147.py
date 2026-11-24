#!/usr/bin/env python3
"""Fee rates that change by date (v1.47): '1.6%; >2026-11-22 1.2%' in the
Deposit / Withdrawal cell.

The period is picked with the row's REPORT date (tentukan_tanggal: Paid Date for
deposits, Completed Date for DPM withdrawals and rebates), so a withdrawal applied
on the 22nd and completed on the 23rd uses the rate from the 23rd.
Plain rates and v1.38 amount tiers must parse exactly as in v1.46.
All data here is dummy.

Run: python3 test_dated_fee_v147.py
"""
import datetime
import importlib.util
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

import hitung_dw as h

ROOT = Path(__file__).resolve().parent
D = datetime.date


def near(a, b, eps=1e-6):
    return a is not None and abs(a - b) < eps


def fee(teks, jumlah, tanggal):
    """Handling Fee the way process_sheet computes it (no legacy column)."""
    k = h.baca_rate(teks)
    assert k is not None and not k.get("error"), (teks, k)
    k, periode = h.pilih_periode(k, tanggal)
    if k.get("tiers"):
        k, _label = h.pilih_tier(k, jumlah, periode)
    f = jumlah * k["pct"]
    if k["min"] is not None:
        f = max(f, k["min"])
    return f + k["fixed"]


# ------------------------------------------------------------------ parsing
k = h.baca_rate("1.6%; >2026-11-22 1.2%")
assert not k.get("error"), k
assert near(k["pct"], 0.016) and k["fixed"] == 0 and k["min"] is None
assert len(k["periode"]) == 1 and k["periode"][0]["mulai"] == D(2026, 11, 23)
assert near(k["periode"][0]["pct"], 0.012)
assert k["teks"] == "1.6%; >2026-11-22 1.2%"

# '>' = strictly after; '>=', '≥' and 'from' = from that day; ';' or '|' before a date
for t in ("1.6%; >2026-11-22 1.2%", "1.6%; >=2026-11-23 1.2%", "1.6%; ≥2026-11-23 1.2%",
          "1.6%; from 2026-11-23 1.2%", "1.6% ; FROM 2026-11-23 1.2%",
          "1.6% | >=2026-11-23 1.2%", "1.6%;>2026-11-22 1.2%;", "1.6%; >2026-11-22   1.2%"):
    assert near(fee(t, 1_000_000, D(2026, 11, 1)), 16_000), t
    assert near(fee(t, 1_000_000, D(2026, 11, 22)), 16_000), t
    assert near(fee(t, 1_000_000, D(2026, 11, 23)), 12_000), t
    assert near(fee(t, 1_000_000, D(2026, 12, 31)), 12_000), t
    assert near(fee(t, 1_000_000, datetime.datetime(2026, 11, 23, 0, 33)), 12_000), t

# several periods; each may use +fixed / min
t = "1.5%+50; >2026-11-10 1%,min 5; >=2026-11-21 +8"
assert near(fee(t, 1000, D(2026, 11, 10)), 65)          # 1.5% + 50
assert near(fee(t, 100, D(2026, 11, 11)), 5)            # 1% = 1 -> min 5
assert near(fee(t, 1000, D(2026, 11, 20)), 10)
assert near(fee(t, 1000, D(2026, 11, 21)), 8)           # flat 8
k = h.baca_rate(t)
assert [p["mulai"] for p in k["periode"]] == [D(2026, 11, 11), D(2026, 11, 21)]

# date periods with their own amount tiers: '|' between periods, ';' between tiers
for t in ("3.5%; >=5000 3.8% | >=2026-11-23 3%; >=5000 3.2%",
          "3.5%; >=5000 3.8%; >=2026-11-23 3%; >=5k 3.2%"):       # same meaning
    assert near(fee(t, 4000, D(2026, 11, 22)), 140), t
    assert near(fee(t, 5000, D(2026, 11, 22)), 190), t
    assert near(fee(t, 4000, D(2026, 11, 23)), 120), t
    assert near(fee(t, 5000, D(2026, 11, 23)), 160), t
# a period without tiers after a tiered one is flat
t = "3.5%; >5000 3.8% | >2026-11-22 3%"
assert near(fee(t, 6000, D(2026, 11, 22)), 228)
assert near(fee(t, 6000, D(2026, 11, 23)), 180)
assert "tiers" not in h.pilih_periode(h.baca_rate(t), D(2026, 11, 23))[0]

# labels (Data Status)
k = h.baca_rate("1.6%; >2026-11-22 1.2%; >=2026-12-01 1%")
assert h.pilih_periode(k, D(2026, 11, 2))[1] == "until 2026-11-22"
assert h.pilih_periode(k, D(2026, 11, 23))[1] == "2026-11-23 to 2026-11-30"
assert h.pilih_periode(k, D(2026, 12, 1))[1] == "from 2026-12-01"
kt, per = h.pilih_periode(h.baca_rate("3.5%; >5000 3.8% | >2026-11-22 3%; >5000 3.2%"),
                          D(2026, 11, 23))
assert h.pilih_tier(kt, 6000, per)[1] == \
    "3.2% (rate from 2026-11-23; tier >5,000 applied to the whole amount)"
assert h.pilih_tier(kt, 10, per)[1].startswith("3% (rate from 2026-11-23; normal rate")
# without a period the v1.38 label is unchanged
assert h.pilih_tier(h.baca_rate("3.5%; >5000 3.8%"), 6000)[1] == \
    "3.8% (tier >5,000 applied to the whole amount)"
assert h.pilih_periode(h.baca_rate("3.5%"), D(2026, 11, 23)) == (h.baca_rate("3.5%"), None)

# ------------------------------------------------------------- bad input
for t, sebab in (
        (">=2026-11-23 1.2%", "rate used before the first date must come first"),
        ("1.6% >2026-11-22 1.2%", "start each date period with ';' or '|'"),
        ("1.6%; >2026-11-22", "no rate after >2026-11-22"),
        ("1.6%; >2026-11-22 abc", "cannot read the rate 'abc' of >2026-11-22"),
        ("1.6%; >22/11/2026 1.2%", "write dates as YYYY-MM-DD"),
        ("1.6%; >22-11-2026 1.2%", "write dates as YYYY-MM-DD"),
        ("1.6%; >2026/11/22 1.2%", "write dates as YYYY-MM-DD"),
        ("1.6%; >2026-02-30 1.2%", "'2026-02-30' is not a real date"),
        ("1.6%; >2026-13-01 1.2%", "is not a real date"),
        ("1.6%; >2026-11-22 1.2%; >=2026-11-23 1%", "both start on 2026-11-23"),
        ("1.6%; >=2026-11-23 1.2%; from 2026-11-23 1%", "both start on 2026-11-23"),
        ("1.6%; >2026-11-25 1.2%; >2026-11-20 1%", "overlap or are out of order"),
        ("1.6%;; >2026-11-22 1.2%", "empty part"),
        ("1.6%; >2026-11-22 1.2% >5000 1%", "separate each tier with ';'"),
        ("1.6% | 1.2% | >2026-11-22 1%", "after '|' a new date period must start"),
        ("1.6%; >2026-11-22 1.2%; >5000 1.4%; >5k 1.5%", "is used by two tiers"),
        ("abc; >2026-11-22 1.2%", "the first period"),
        ("1.6%; >2026-11-221.2%", "start each date period")):
    r = h.baca_rate(t)
    assert r.get("error") and sebab in r["error"], (t, r)
    assert r["pct"] == 0 and r["teks"] == " ".join(t.split()), r
    note = h.catatan_tier({"deposit": r})
    assert note and "cannot be read" in note[0] and "EMPTY" in note[0], note

# ------------------------------------------- existing formats unchanged
# Compare with the v1.46 parser (commit 6bbc4ad) when git history is available.
LAMA = None
_git = subprocess.run(["git", "show", "6bbc4ad:hitung_dw.py"], cwd=ROOT,
                      capture_output=True, text=True)
if _git.returncode == 0:
    _p = Path(tempfile.mkdtemp()) / "hitung_dw_v146.py"
    _p.write_text(_git.stdout)
    spec = importlib.util.spec_from_file_location("hitung_dw_v146", _p)
    LAMA = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(LAMA)
    print("compared with v1.46 baca_rate")
KORPUS = [None, "", " ", 0, 0.035, 1, 8, 2.7, "0.027", "2.7%", "2,7%", "+8", "+ 8", "1.5%+50",
          "0.45%,min 90", "0.45%, MIN 90", "min 90", "8", "8%", "1.234,5", "1,234", "abc",
          "3.5%;", "1.5%; min 50", "2%+1,000", "0.45%,minimum 90", "3%+", "(1)", "1.5% + 50",
          "1.5%+50 min 10", "2 %", "2026", "2026-11-22", "1.6% from Nov", "1.6% | 2%",
          "3.5%; >5000 3.8%", "3.5%; >=5000 3.8%", "3.5% ; ≥5000 3.8%",
          "3.5%; >5k 3.8%; >20k 4%", "3.5%; >20k 4%; >5k 3.8%", "1.5%+50; >5000 1.2%+50",
          "3.5%; >5,000 3.8%", "3.5%; >5.000 3.8%", "3.5%; >1.000.000 3%", "3.5%; >2026 3%",
          "0.45%,min 90; >10000 0.4%,min 80", "3.5% >5000 3.8%", ">5000 3.8%",
          "3.5%; >abc 3.8%", "3.5%; >5000", "3.5%; >5000 3.8%; >5k 4%"]
if LAMA is not None:
    for v in KORPUS:
        assert h.baca_rate(v) == LAMA.baca_rate(v), (v, h.baca_rate(v), LAMA.baca_rate(v))
for v in KORPUS:
    r = h.baca_rate(v)
    assert r is None or "periode" not in r, (v, r)

DEP0 = "1.6%; >2026-11-22 1.2%"
# fee-sheet cells keep dated rates as text (re-readable) and the Note explains them
t = "1.6%; >2026-11-22 1.2%"
assert h.sel_rate(h.baca_rate(t)) == (t, None, None)
assert h.catatan_tier({"deposit": h.baca_rate(t)}) == [
    "Deposit rate by date: until 2026-11-22: 1.6% | from 2026-11-23: 1.2% "
    "(date = the row's report date)"]
n = h.catatan_tier({"withdrawal": h.baca_rate(
    "3.5%; >5000 3.8% | >=2026-11-23 3%; >5000 3.2%")})[0]
assert n == ("Withdrawal rate by date: until 2026-11-22: 3.5%; >5,000 3.8% | "
             "from 2026-11-23: 3%; >5,000 3.2% (tier rate on the whole amount) "
             "(date = the row's report date)"), n
# an informative note is not a Missing Data warning
assert h.kumpulkan_peringatan_tier({("VND", "X"): {"deposit": h.baca_rate(t), "nama": "X"}}) == []
# a Remark that looks like a tier rule is fine when the dated cell has tiers
assert not any("Remark looks like" in x for x in h.catatan_tier(
    {"deposit": h.baca_rate("1% | >2026-11-22 1%; >5000 2%"), "remark": "more than 5000 then 2%"}))


# a period rate that reads as 100% or more is named like a plain cell
w = h.kumpulkan_rate_tinggi({("VND", "X"): {"deposit": h.baca_rate("1.6%; >2026-11-22 8"),
                                            "nama": "X"}})
assert w == ["VND-X deposit rate 8 = 800% (period >2026-11-22). If this is a flat fee per "
             "transaction, write +8."], w
assert h.kumpulkan_rate_tinggi({("VND", "X"): {"deposit": h.baca_rate(DEP0), "nama": "X"}}) == []

# --------------------------------------- end to end, DPM layout
def _detail(book):
    rows = list(book["D&W Detail"].iter_rows(values_only=True))
    hd = list(rows[0])
    return {r[hd.index("Ticket")]: dict(zip(hd, r)) for r in rows[1:]
            if r and r[hd.index("Ticket")]}


def _teks(ws):
    return "\n".join(str(v) for row in ws.iter_rows(values_only=True) for v in row
                     if v is not None)


def _jalan(*args):
    r = subprocess.run([sys.executable, *map(str, args)], cwd=ROOT,
                       capture_output=True, text=True)
    assert r.returncode == 0, (args[0], r.stdout[-3000:], r.stderr[-3000:])
    return r.stdout


DEP = "1.6%; >2026-11-22 1.2%"
WD = "3.5%; >=5000 3.8% | >=2026-11-23 3%; >=5000 3.2%"
N22, N23 = datetime.datetime(2026, 11, 22), datetime.datetime(2026, 11, 23)
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, dst = tmp / "dpm.xlsx", tmp / "out.xlsx"
    wb = Workbook()
    d = wb.active
    d.title = "D"
    d.append(["Source Name", "Currency", "Paid Date", "Transaction", "Payment Gateway",
              "USD", "Rate", "Ticket"])
    d.append(["Deposit", "VND", N22, 1_000_000, "BeckPay-VND", 40, 25000, "D22"])
    d.append(["Deposit", "VND", N23, 1_000_000, "BeckPay-VND", 40, 25000, "D23"])
    d.append(["Deposit", "VND", N23, 1_000_000, "BadDate", 40, 25000, "DBAD"])
    d.append(["Deposit", "VND", N23, 1_000_000, "Plain", 40, 25000, "DPLAIN"])
    w = wb.create_sheet("W")
    w.append(["Source Name", "Currency", "Apply Date", "Settlement Date", "Completed Date",
              "Transaction", "Payment Gateway", "USD", "Rate", "Ticket", "Status"])
    # applied / settled on the 22nd, completed on the 23rd -> the 23rd's rate
    w.append(["Withdrawal", "VND", N22, N22, N23, 6000, "BeckPay-VND", 0.24, 25000, "W1",
              "finish"])
    w.append(["Withdrawal", "VND", N22, N22, N22, 6000, "BeckPay-VND", 0.24, 25000, "W2",
              "finish"])
    w.append(["Withdrawal", "VND", N22, N22, N23, 4000, "BeckPay-VND", 0.16, 25000, "W3",
              "success"])
    w.append(["Rebate Withdrawal", "VND", N22, N22, N23, 6000, "BeckPay-VND", 0.24, 25000,
              "R1", "finish"])
    x = wb.create_sheet("Xero")
    x.append(["Date", "VND", "USD"])
    for day in range(1, 31):
        x.append([datetime.datetime(2026, 11, day), 25000, 1])
    f = wb.create_sheet("D&W FEE")
    f.append(["Currency", "Payment Gateway", "Deposit", "Withdrawal", "Remark"])
    f.append(["VND", "BeckPay-VND", DEP, WD, None])
    f.append(["VND", "BadDate", "1.6%; >2026-11-31 1.2%", 0.01, None])
    f.append(["VND", "Plain", 0.016, 0.01, None])
    ft = wb.create_sheet("Fund Transfer Table")
    ft.append(["Date", "Department", "Payment Gateway", "Currency", "Amount", "Fee",
               "Xero (USD)", "收款日期", "收款部门", "收款人", "收入币种", "金额"])
    wb.save(src)
    _jalan("hitung_dw.py", src, "-o", dst, "--period", "2026-11", "--tanpa-extra-fees")
    book = load_workbook(dst, data_only=True)
    det = _detail(book)
    # deposit: 22 Nov 1.6%, 23 Nov 1.2%
    assert near(det["D22"]["Handling Fee"], 16_000), det["D22"]
    assert near(det["D23"]["Handling Fee"], 12_000), det["D23"]
    assert det["D22"]["Data Status"] == "OK; FEE: 1.6% (rate until 2026-11-22)", det["D22"]
    assert det["D23"]["Data Status"] == "OK; FEE: 1.2% (rate from 2026-11-23)", det["D23"]
    # withdrawal: report date = Completed Date, not Apply / Settlement Date
    assert det["W1"]["Date"].date() == D(2026, 11, 23)
    assert near(det["W1"]["Handling Fee"], 6000 * 0.032), det["W1"]
    assert near(det["W2"]["Handling Fee"], 6000 * 0.038), det["W2"]
    assert near(det["W3"]["Handling Fee"], 4000 * 0.03), det["W3"]
    assert near(det["R1"]["Handling Fee"], 6000 * 0.032), det["R1"]
    assert det["W1"]["Data Status"] == \
        "OK; FEE: 3.2% (rate from 2026-11-23; tier >=5,000 applied to the whole amount)"
    assert det["W2"]["Data Status"] == \
        "OK; FEE: 3.8% (rate until 2026-11-22; tier >=5,000 applied to the whole amount)"
    # bad date: empty (red) fee + clear note, never a silent 1.6%
    assert det["DBAD"]["Handling Fee"] is None, det["DBAD"]
    assert "'2026-11-31' is not a real date" in det["DBAD"]["Data Status"], det["DBAD"]
    assert near(det["DPLAIN"]["Handling Fee"], 16_000) and det["DPLAIN"]["Data Status"] == "OK"
    hilang = _teks(book["Missing Data"])
    assert "'2026-11-31' is not a real date" in hilang
    assert "Rate by date: 1.6%; >2026-11-22 1.2%" in hilang
    fs = book[h.FEE_SHEET]
    rows = {r[1]: r for r in fs.iter_rows(min_row=4, values_only=True) if r[1]}
    assert rows["BECKPAY-VND"][2] == DEP and rows["BECKPAY-VND"][3] == WD, rows["BECKPAY-VND"]
    assert "Deposit rate by date: until 2026-11-22: 1.6% | from 2026-11-23: 1.2%" \
        in rows["BECKPAY-VND"][9]
    assert "Withdrawal rate by date" in rows["BECKPAY-VND"][9]
    book.close()
    # the written fee sheet reads back with the same periods
    tbl = h.load_fee_table(load_workbook(dst, data_only=True))
    v = next(v for v in tbl.values() if v.get("nama") == "BECKPAY-VND")
    assert v["deposit"] == h.baca_rate(DEP), v["deposit"]
    assert v["withdrawal"] == h.baca_rate(WD), v["withdrawal"]

# separate 'Deposit Fixed' / 'Deposit Min' columns also apply to date periods
wb = Workbook()
ws = wb.active
ws.title = "D&W FEE"
ws.append(["Currency", "Payment Gateway", "Deposit", "Withdrawal", "Deposit Fixed",
           "Deposit Min"])
ws.append(["VND", "X", "1% | >2026-11-22 2%; >5000 3%", 0.01, 5, 7])
ws.append(["VND", "Y", "1% | >2026-11-22 2%+1 min 2", 0.01, 5, 7])
tbl = h.load_fee_table(wb, pakai_extra_fees=False)
kx = tbl[("VND", h.kunci_gw("X"))]["deposit"]
assert (kx["fixed"], kx["min"]) == (5, 7)
assert (kx["periode"][0]["fixed"], kx["periode"][0]["min"]) == (5, 7)
assert (kx["periode"][0]["tiers"][0]["fixed"], kx["periode"][0]["tiers"][0]["min"]) == (5, 7)
ky = tbl[("VND", h.kunci_gw("Y"))]["deposit"]
assert (ky["periode"][0]["fixed"], ky["periode"][0]["min"]) == (1, 2)   # its own wins

# a row with no date cannot pick a period -> empty fee + note (no --period filter)
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, dst = tmp / "nodate.xlsx", tmp / "out.xlsx"
    wb = Workbook()
    d = wb.active
    d.title = "D"
    d.append(["Source Name", "Currency", "Paid Date", "Transaction", "Payment Gateway",
              "USD", "Rate", "Ticket"])
    d.append(["Deposit", "VND", N23, 1_000_000, "BeckPay-VND", 40, 25000, "OK1"])
    d.append(["Deposit", "VND", None, 1_000_000, "BeckPay-VND", 40, 25000, "ND1"])
    x = wb.create_sheet("Xero")
    x.append(["Date", "VND"])
    x.append([N23, 25000])
    f = wb.create_sheet("D&W FEE")
    f.append(["Currency", "Payment Gateway", "Deposit", "Withdrawal"])
    f.append(["VND", "BeckPay-VND", DEP, 0.01])
    wb.save(src)
    _jalan("hitung_dw.py", src, "-o", dst, "--tanpa-extra-fees")
    book = load_workbook(dst, data_only=True)
    det = _detail(book)
    assert near(det["OK1"]["Handling Fee"], 12_000), det["OK1"]
    assert det["ND1"]["Handling Fee"] is None, det["ND1"]
    assert "changes by date but this row has no date" in det["ND1"]["Data Status"], det["ND1"]
    assert "changes by date but these rows have no date" in _teks(book["Missing Data"])
    book.close()

# -------------------------------------- end to end, KVB upload template
import app as app_mod  # noqa: E402

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    wb = load_workbook(ROOT / "Template KVB Plus.xlsx")
    d = wb["D"]
    hd = [c.value for c in d[1]]
    for r, (tgl, t) in enumerate(((N22, "K22"), (N23, "K23")), start=2):
        vals = {"Source.Name": "11-2026 Deposit", "Paid Date": tgl,
                "Settlement Date": tgl, "method": "VND", "ticket": t,
                "USD (CRM)": 40, "rate": 25000, "transaction": 1_000_000,
                "receive_company_account": "BeckPay-VND"}
        for kk, v in vals.items():
            d.cell(r, hd.index(kk) + 1, v)
    x = wb["Xero"]
    hx = [c.value for c in x[1]]
    for day in range(1, 31):
        x.cell(day + 1, 1, datetime.datetime(2026, 11, day))
        for cur, val in (("VND", 25000), ("USD", 1), ("USDT", 1)):
            x.cell(day + 1, hx.index(cur) + 1, val)
    fee_ws = wb["Handling Fee"]
    hf = [c.value for c in fee_ws[1]]
    for kk, v in {"Payment Channel": "BeckPay-VND", "Currency": "VND",
                  "Deposit": DEP, "Withdrawal": 0.01}.items():
        fee_ws.cell(2, hf.index(kk) + 1, v)
    wb.save(tmp / "kvb.xlsx")
    _jalan("siapkan_kvb.py", tmp / "kvb.xlsx", "-o", tmp / "t.xlsx")
    _jalan("buat_template.py", tmp / "t.xlsx", "-o", tmp / "tpl0.xlsx")
    _jalan("isi_template.py", tmp / "t.xlsx", "-t", tmp / "tpl0.xlsx", "--bulan", "2026-11",
           *app_mod.KVB_ARG["isi_template.py"], "-o", tmp / "tpl.xlsx")
    _jalan("hitung_dw.py", tmp / "tpl.xlsx", "--period", "2026-11", "-o", tmp / "hasil.xlsx",
           *app_mod.KVB_ARG["hitung_dw.py"])
    book = load_workbook(tmp / "hasil.xlsx", data_only=True)
    det = _detail(book)
    assert near(det["K22"]["Handling Fee"], 16_000), det["K22"]
    assert near(det["K23"]["Handling Fee"], 12_000), det["K23"]
    assert "rate from 2026-11-23" in det["K23"]["Data Status"], det["K23"]
    book.close()

# ------------------------------------------------------- guides
import buat_template_kvb as btk  # noqa: E402

assert any(a == "1.6%; >2026-11-22 1.2%" for a, _ in btk.FEE_CARA_ISI)
guide = _teks(load_workbook(ROOT / "Template KVB Plus.xlsx")["Guide"])
assert "1.6%; >2026-11-22 1.2%" in guide and "YYYY-MM-DD" in guide
assert "3.5%; >5000 3.8% | >=2026-11-23 3%; >5000 3.2%" in guide
box = [c.value for c in load_workbook(ROOT / "Template KVB Plus.xlsx")["Handling Fee"]["I"]]
assert "1.6%; >2026-11-22 1.2%" in box
assert "Rate by date" in (ROOT / "buat_template.py").read_text()
assert "&gt;2026-11-22 1.2%" in (ROOT / "app.py").read_text()
assert "&gt;2026-11-22 1.2%" in (ROOT / "templates" / "tool.html").read_text()
assert "Rate by date '1.6%; >2026-11-22 1.2%'" in h.LEGENDA["Handling Fee"]

print("test_dated_fee_v147: ok")
