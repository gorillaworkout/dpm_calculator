#!/usr/bin/env python3
"""Tiered fee rates (v1.38): '3.5%; >5000 3.8%' in the Deposit / Withdrawal cell.

The tier rate applies to the WHOLE amount when the transaction (in its own
currency) passes the threshold. All data here is dummy.

Run: python3 test_tiered_fee.py
"""
import datetime
import importlib.util
import io
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

import hitung_dw as h

ROOT = Path(__file__).resolve().parent


def near(a, b, eps=1e-9):
    return a is not None and abs(a - b) < eps


def fee(teks, jumlah):
    """Handling Fee the way process_sheet computes it (no legacy column)."""
    k = h.baca_rate(teks)
    assert k is not None and not k.get("error"), (teks, k)
    k, _label = h.pilih_tier(k, jumlah)
    f = jumlah * k["pct"]
    if k["min"] is not None:
        f = max(f, k["min"])
    return f + k["fixed"]


# ------------------------------------------------------------------ parsing
k = h.baca_rate("3.5%; >5000 3.8%")
assert near(k["pct"], 0.035) and k["fixed"] == 0 and k["min"] is None
assert len(k["tiers"]) == 1 and k["tiers"][0]["op"] == ">" and k["tiers"][0]["batas"] == 5000
assert near(k["tiers"][0]["pct"], 0.038)

# below / equal / above (strict '>')
assert near(fee("3.5%; >5000 3.8%", 4999), 4999 * 0.035)
assert near(fee("3.5%; >5000 3.8%", 5000), 5000 * 0.035)
assert near(fee("3.5%; >5000 3.8%", 5000.01), 5000.01 * 0.038)
assert near(fee("3.5%; >5000 3.8%", 6000), 228.0)              # whole amount, not marginal
# '>=' is inclusive
assert near(fee("3.5%; >=5000 3.8%", 5000), 190.0)
assert near(fee("3.5%; >=5000 3.8%", 4999.99), 4999.99 * 0.035)
assert near(fee("3.5% ; ≥5000 3.8%", 5000), 190.0)
# several tiers, any order: highest threshold passed wins
for t in ("3.5%; >5k 3.8%; >20k 4%", "3.5%; >20k 4%; >5k 3.8%"):
    assert near(fee(t, 5000), 175.0), t
    assert near(fee(t, 10000), 380.0), t
    assert near(fee(t, 20000), 760.0), t
    assert near(fee(t, 20001), 800.04), t
# k / K shorthand and thousands separators
for t in ("3.5%; >5k 3.8%", "3.5%; >5K 3.8%", "3.5%; >5,000 3.8%", "3.5%; >5.000 3.8%",
          "3.5%;>5000 3.8%", "3.5%;  > 5000   3.8%", "3.5%; >5000 3.8%;"):
    assert h.baca_rate(t)["tiers"][0]["batas"] == 5000, t
assert h.baca_rate("1%; >1.5k 2%")["tiers"][0]["batas"] == 1500
assert h.baca_rate("1%; >1,000,000 0.5%")["tiers"][0]["batas"] == 1_000_000
# combined with +fixed and min, per tier
assert near(fee("1.5%+50; >5000 1.2%+50", 1000), 65.0)
assert near(fee("1.5%+50; >5000 1.2%+50", 10000), 170.0)
assert near(fee("0.45%,min 90; >10000 0.4%", 1000), 90.0)        # min on the normal rate
assert near(fee("0.45%,min 90; >10000 0.4%", 30000), 120.0)      # tier has no min
assert near(fee("0.45%,min 90; >10000 0.4%,min 50", 11000), 50.0)
assert near(fee("+8; >1000 +12", 500), 8.0) and near(fee("+8; >1000 +12", 1500), 12.0)
# negative amounts (withdrawal exports) are compared by size
assert near(fee("3.5%; >5000 3.8%", -6000), -228.0)

# labels used in D&W Detail 'Data Status'
_k, lab = h.pilih_tier(h.baca_rate("3.5%; >5000 3.8%"), 6000)
assert lab == "3.8% (tier >5,000 applied to the whole amount)", lab
_k, lab = h.pilih_tier(h.baca_rate("3.5%; >5000 3.8%"), 5000)
assert lab == "3.5% (normal rate; tier from >5,000 not reached)", lab
assert h.pilih_tier(h.baca_rate("3.5%"), 6000)[1] is None

# invalid syntax -> error dict, never a silently wrong number
for t, sebab in (("3.5% >5000 3.8%", "separate each tier"),
                 ("3.5%; >abc 3.8%", "cannot read tier"),
                 ("3.5%; >5000", "cannot read tier"),
                 ("3.5%; >5000 3.8%", None),
                 ("; >5000 3%", "normal rate must come first"),
                 (">5000 3.8%", "normal rate must come first"),
                 ("3.5%; 5000 3.8%; >1 2%", "cannot read tier"),
                 ("3.5%; >5000 3.8%; >5k 4%", "is used by two tiers"),
                 ("abc; >5000 3.8%", "cannot read the normal rate"),
                 ("3.5%; >5000 abc", "cannot read tier")):
    r = h.baca_rate(t)
    if sebab is None:
        assert not r.get("error"), (t, r)
    else:
        assert r.get("error") and sebab in r["error"], (t, r)
        assert r["teks"] == t.strip(), r

# ------------------------------------------- existing formats byte-for-byte
# Compare with the v1.37 parser (commit c6a7385) when git history is available.
LAMA = None
_git = subprocess.run(["git", "show", "c6a7385:hitung_dw.py"], cwd=ROOT,
                      capture_output=True, text=True)
if _git.returncode == 0:
    _tmpd = tempfile.mkdtemp()
    _p = Path(_tmpd) / "hitung_dw_v137.py"
    _p.write_text(_git.stdout)
    sys.path.insert(0, str(ROOT))
    spec = importlib.util.spec_from_file_location("hitung_dw_v137", _p)
    LAMA = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(LAMA)
    print("compared with v1.37 baca_rate")
KORPUS = [None, "", " ", 0, 0.035, 1, 8, 2.7, "0.027", "2.7%", "2,7%", "+8", "+ 8", "1.5%+50",
          "0.45%,min 90", "0.45%, MIN 90", "min 90", "8", "8%", "1.234,5", "1,234", "abc",
          "3.5%;", "1.5%; min 50", "2%+1,000", "0.45%,minimum 90", "3%+", "(1)", "1.5% + 50",
          "1.5%+50 min 10", "2 %"]
if LAMA is not None:
    for v in KORPUS:
        assert h.baca_rate(v) == LAMA.baca_rate(v), (v, h.baca_rate(v), LAMA.baca_rate(v))
for v in KORPUS:                               # no tier -> no tier keys, ever
    r = h.baca_rate(v)
    assert r is None or set(r) == {"pct", "fixed", "min"}, (v, r)

# ------------------------------------------------------------ Remark hints
assert h.remark_mirip_tier("for Deposit transactions more than 5K, then 3.8%")
assert h.remark_mirip_tier("lebih dari 5.000 pakai 3,8%")
assert h.remark_mirip_tier("above 10000 = 0.4%")
assert h.remark_mirip_tier("> 5000 3.8%")
assert not h.remark_mirip_tier("ACTIVE")
assert not h.remark_mirip_tier("rate changed on 2026-07-01")
assert not h.remark_mirip_tier(None)

# fee-sheet cells keep tiered rates as text (re-readable)
assert h.sel_rate(h.baca_rate("3.5%; >5000 3.8%")) == ("3.5%; >5000 3.8%", None, None)
assert h.sel_rate(h.baca_rate("1.5%+50")) == (0.015, 50.0, None)
assert h.sel_rate(None) == (None, None, None)


# -------------------------------------------- end to end, DPM layout
def _sheets(wb, fee_rows, fee_header):
    d = wb.active
    d.title = "D"
    d.append(["Source Name", "Currency", "Paid Date", "Transaction", "Payment Gateway",
              "USD", "Rate", "Ticket"])
    for i, tx in enumerate((4000, 5000, 6000, 20000), start=1):
        d.append(["Deposit", "USD", datetime.datetime(2026, 7, 10), tx, "Pay247-KH USD",
                  tx, 1, f"D{i}"])
    d.append(["Deposit", "USD", datetime.datetime(2026, 7, 10), 9000, "BadTier", 9000, 1, "D9"])
    d.append(["Deposit", "USD", datetime.datetime(2026, 7, 10), 9000, "RemarkOnly", 9000, 1,
              "D10"])
    w = wb.create_sheet("W")
    w.append(["Source Name", "Currency", "Completed Date", "Transaction", "Payment Gateway",
              "USD", "Rate", "Ticket", "Status"])
    w.append(["Withdrawal", "USD", datetime.datetime(2026, 7, 11), 8000, "Pay247-KH USD",
              8000, 1, "W1", "finish"])
    x = wb.create_sheet("Xero")
    x.append(["Date", "USD"])
    for day in range(1, 32):
        x.append([datetime.datetime(2026, 7, day), 1])
    f = wb.create_sheet("D&W FEE")
    f.append(fee_header)
    for r in fee_rows:
        f.append(r)
    ft = wb.create_sheet("Fund Transfer Table")
    ft.append(["Date", "Department", "Payment Gateway", "Currency", "Amount", "Fee",
               "Xero (USD)", "收款日期", "收款部门", "收款人", "收入币种", "金额"])


def _detail(book):
    rows = list(book["D&W Detail"].iter_rows(values_only=True))
    hd = list(rows[0])
    out = {}
    for r in rows[1:]:
        if r and r[hd.index("Ticket")]:
            out[r[hd.index("Ticket")]] = dict(zip(hd, r))
    return out


def _teks(ws):
    return "\n".join(str(v) for row in ws.iter_rows(values_only=True) for v in row
                     if v is not None)


REMARK = "for Deposit transactions more than 5K, then 3.8%"
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, dst = tmp / "dpm.xlsx", tmp / "out.xlsx"
    wb = Workbook()
    _sheets(wb, [["USD", "Pay247-KH USD", "3.5%; >5000 3.8%", 0.02, "tiered example"],
                 ["USD", "BadTier", "3.5% >5000 3.8%", 0.02, None],
                 ["USD", "RemarkOnly", 0.035, 0.02, REMARK]],
            ["Currency", "Payment Gateway", "Deposit", "Withdrawal", "Remark"])
    wb.save(src)
    proc = subprocess.run([sys.executable, str(ROOT / "hitung_dw.py"), str(src), "-o", str(dst),
                           "--period", "2026-07", "--tanpa-extra-fees"],
                          cwd=ROOT, capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr[-3000:] + proc.stdout[-3000:]
    book = load_workbook(dst, data_only=True)
    det = _detail(book)
    assert near(det["D1"]["Handling Fee"], 140.0, 1e-6), det["D1"]
    assert near(det["D2"]["Handling Fee"], 175.0, 1e-6), det["D2"]       # 5,000 is not > 5,000
    assert near(det["D3"]["Handling Fee"], 228.0, 1e-6), det["D3"]
    assert near(det["D4"]["Handling Fee"], 760.0, 1e-6), det["D4"]
    assert near(det["W1"]["Handling Fee"], 160.0, 1e-6), det["W1"]       # withdrawal not tiered
    assert det["D3"]["Data Status"] == "OK; FEE: 3.8% (tier >5,000 applied to the whole amount)"
    assert det["D1"]["Data Status"].startswith("OK; FEE: 3.5% (normal rate")
    assert det["W1"]["Data Status"] == "OK"
    # invalid tier: no fee (red), clear note, not 3.5% silently
    assert det["D9"]["Handling Fee"] is None, det["D9"]
    assert "cannot be read" in det["D9"]["Data Status"], det["D9"]
    # remark-only tier: calculated at the cell rate, plus a warning
    assert near(det["D10"]["Handling Fee"], 315.0, 1e-6), det["D10"]
    hilang = _teks(book["Missing Data"])
    assert "2c. TIERED FEE RATES TO CHECK" in hilang
    assert "Remark looks like a tier rule" in hilang and REMARK in hilang
    assert "separate each tier with ';'" in hilang
    fs = book[h.FEE_SHEET]
    rows = {r[1]: r for r in fs.iter_rows(min_row=4, values_only=True) if r[1]}
    assert rows["PAY247-KH USD"][2] == "3.5%; >5000 3.8%", rows["PAY247-KH USD"]
    assert "Deposit tiered" in rows["PAY247-KH USD"][9]
    assert "Remark looks like a tier rule" in rows["REMARKONLY"][9]
    book.close()
    # the written fee sheet reads back with the same tiers
    tbl = h.load_fee_table(load_workbook(dst, data_only=True))
    k2 = next(v for v in tbl.values() if v.get("nama") == "PAY247-KH USD")["deposit"]
    assert k2["tiers"] == h.baca_rate("3.5%; >5000 3.8%")["tiers"]

# -------------------------------------- end to end, KVB upload template
import app as app_mod  # noqa: E402

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    wb = load_workbook(ROOT / "Template KVB Plus.xlsx")
    d = wb["D"]
    hd = [c.value for c in d[1]]
    for r, (tx, t) in enumerate(((4000, "K1"), (6000, "K2")), start=2):
        vals = {"Source.Name": "07-2026 Deposit", "Paid Date": datetime.datetime(2026, 7, 3),
                "Settlement Date": datetime.datetime(2026, 7, 3), "method": "USD",
                "ticket": t, "USD (CRM)": tx, "rate": 1, "transaction": tx,
                "receive_company_account": "Pay247-KH USD"}
        for kk, v in vals.items():
            d.cell(r, hd.index(kk) + 1, v)
    x = wb["Xero"]
    hx = [c.value for c in x[1]]
    for day in range(1, 32):
        x.cell(day + 1, 1, datetime.datetime(2026, 7, day))
        for cur in ("USD", "USDT"):
            x.cell(day + 1, hx.index(cur) + 1, 1)
    fee_ws = wb["Handling Fee"]
    hf = [c.value for c in fee_ws[1]]
    for kk, v in {"Payment Channel": "Pay247-KH USD", "Currency": "USD",
                  "Deposit": "3.5%; >5000 3.8%", "Withdrawal": 0.02,
                  "Remark": REMARK}.items():
        fee_ws.cell(2, hf.index(kk) + 1, v)
    wb.save(tmp / "kvb.xlsx")

    def jalan(*args):
        r = subprocess.run([sys.executable, *map(str, args)], cwd=ROOT,
                           capture_output=True, text=True)
        assert r.returncode == 0, (args[0], r.stdout[-3000:], r.stderr[-3000:])
        return r.stdout

    jalan("siapkan_kvb.py", tmp / "kvb.xlsx", "-o", tmp / "t.xlsx")
    jalan("buat_template.py", tmp / "t.xlsx", "-o", tmp / "tpl0.xlsx")
    jalan("isi_template.py", tmp / "t.xlsx", "-t", tmp / "tpl0.xlsx", "--bulan", "2026-07",
          *app_mod.KVB_ARG["isi_template.py"], "-o", tmp / "tpl.xlsx")
    jalan("hitung_dw.py", tmp / "tpl.xlsx", "--period", "2026-07", "-o", tmp / "hasil.xlsx",
          *app_mod.KVB_ARG["hitung_dw.py"])
    book = load_workbook(tmp / "hasil.xlsx", data_only=True)
    det = _detail(book)
    assert near(det["K1"]["Handling Fee"], 140.0, 1e-6), det["K1"]
    assert near(det["K2"]["Handling Fee"], 228.0, 1e-6), det["K2"]
    assert "tier >5,000 applied" in det["K2"]["Data Status"], det["K2"]
    book.close()

# ------------------------------------------------------- template guides
import buat_template_kvb as btk  # noqa: E402

assert "3.5%; >5000 3.8%" in btk.FEE_PANDUAN and "WHOLE amount" in btk.FEE_PANDUAN
assert "Remark is not read" in btk.FEE_PANDUAN
guide = _teks(load_workbook(ROOT / "Template KVB Plus.xlsx")["Guide"])
for s in ("3.5%; >5000 3.8%", "3.5%; >=5000 3.8%", "3.5%; >5k 3.8%; >20k 4%",
          "1.5%+50; >5000 1.2%+50", "Remark is NOT read"):
    assert s in guide, s
dpm_src = (ROOT / "buat_template.py").read_text()
assert "3.5%; >5000 3.8%" in dpm_src and "Not in Remark" in dpm_src
app_src = (ROOT / "app.py").read_text()
assert "3.5%; &gt;5000 3.8%" in app_src and "not only in Remark" in app_src
assert io  # noqa

print("test_tiered_fee: ok")
