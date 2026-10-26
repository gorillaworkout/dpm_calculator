"""Withdrawal rows whose 'Completed Date' is blank / text / 1970 are no longer dropped.

Real incident (Sep 2026): ~101,000 Rebate Withdrawal rows in tab W had a blank, text
('success') or 1970 Completed Date but a real datetime in 'Completed Time'. The month
filter in isi_template.py and hitung_dw.py looked at ONE column, so they vanished and the
Missing Data sheet said "None".

v1.29: for Source Name = "Rebate Withdrawal" the 'Completed Date' column is NOT a real
completion date (3,593 rows, all UID 100007, dates 1 Aug-1 Sep 2026 on rows applied 2024-2026),
so it is ignored for Rebate: Completed Time -> Settlement Date. Withdrawal is unchanged.

python3 test_tanggal_fallback.py
"""
import datetime
import importlib.util
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

HERE = Path(__file__).resolve().parent
PY = sys.executable
sys.path.insert(0, str(HERE))
import tanggal_util as tu  # noqa: E402

D = datetime.date
DT = datetime.datetime
SEP = 46266                     # 1 Sep 2026 (Excel serial)
AUG = 46235                     # 1 Aug 2026


def hasil(nilai, utama="COMPLETED DATE"):
    return tu.tentukan_tanggal({k.upper(): v for k, v in nilai.items()}, utama)


# ------------------------------------------------------------------ unit: the helper
# blank Completed Date + datetime Completed Time  -> Completed Time
h = hasil({"Completed Date": None, "Completed Time": SEP + 17.4, "Settlement Date": SEP + 16})
assert h.tanggal == D(2026, 9, 18) and h.sumber == "COMPLETED TIME", h
h = hasil({"Completed Date": None, "Completed Time": DT(2026, 9, 18, 9, 49, 55)})
assert h.tanggal == D(2026, 9, 18) and h.sumber == "COMPLETED TIME", h
# text 'success' in Completed Date (shifted export) -> Completed Time
h = hasil({"Completed Date": "success", "Completed Time": SEP + 3.5, "Settlement Date": SEP + 2})
assert h.tanggal == D(2026, 9, 4) and h.sumber == "COMPLETED TIME", h
# a bare time-of-day in Completed Time is NOT a date
h = hasil({"Completed Date": None, "Completed Time": 0.4097, "Settlement Date": SEP + 2})
assert h.tanggal == D(2026, 9, 3) and h.sumber == "SETTLEMENT DATE", h
h = hasil({"Completed Date": None, "Completed Time": datetime.time(9, 0), "Settlement Date": None})
assert h.tanggal is None and h.alasan_utama == "kosong", h
# Settlement-only
h = hasil({"Completed Date": None, "Completed Time": None, "Settlement Date": DT(2026, 9, 9)})
assert h.tanggal == D(2026, 9, 9) and h.sumber == "SETTLEMENT DATE", h
# 1970: never rescued through Settlement (it means "not paid yet"), only via Completed Time
h = hasil({"Completed Date": DT(1970, 1, 1), "Completed Time": None, "Settlement Date": SEP})
assert h.tanggal is None and h.alasan_utama == "tidak_sah", h
h = hasil({"Completed Date": 25569, "Completed Time": SEP + 5.5, "Settlement Date": SEP + 4})
assert h.tanggal == D(2026, 9, 6) and h.sumber == "COMPLETED TIME", h
# Apply Date is never a fallback
h = hasil({"Completed Date": None, "Completed Time": None, "Settlement Date": None,
           "Apply Date": SEP})
assert h.tanggal is None, h
# explicit valid Completed Date wins; disagreement is FLAGGED, not silently overridden
backfill = {"Completed Date": SEP + 5, "Completed Time": DT(2025, 3, 3, 10, 0),
            "Settlement Date": DT(2025, 3, 3)}
assert tu.UTAMAKAN_WAKTU is False
h = hasil(backfill)
assert h.tanggal == D(2026, 9, 6) and h.sumber == "COMPLETED DATE" and h.konflik, h
tu.UTAMAKAN_WAKTU = True                      # the switch Bayu can flip
try:
    h = hasil(backfill)
    assert h.tanggal == D(2025, 3, 3) and h.konflik, h
finally:
    tu.UTAMAKAN_WAKTU = False
# no conflict when Settlement agrees with Completed Date (Completed Time alone is not enough)
h = hasil({"Completed Date": SEP + 5, "Completed Time": DT(2026, 8, 31, 23, 59),
           "Settlement Date": SEP + 4})
assert h.tanggal == D(2026, 9, 6) and not h.konflik, h
# other primary columns (Paid Date, deposits) are untouched: no fallback chain
h = tu.tentukan_tanggal({"PAID DATE": None, "SETTLEMENT DATE": SEP}, "PAID DATE")
assert h.tanggal is None, h

# ------------------------------------------------------------------ unit: Rebate Withdrawal (v1.29)
def reb(src="Rebate Withdrawal", **kw):
    n = {"Source Name": src}
    n.update(kw)
    return hasil(n)


# Completed Date says August, Completed Time says September -> September (Completed Time)
h = reb(**{"Completed Date": DT(2026, 8, 10), "Completed Time": SEP + 3.5,
           "Settlement Date": SEP + 3})
assert h.tanggal == D(2026, 9, 4) and h.sumber == "COMPLETED TIME" and h.aturan == "rebate", h
assert h.konflik is None and h.selisih_abaikan is True, h
# Completed Date 1 Sep, Completed Time in 2025 -> NOT September
h = reb(**{"Completed Date": SEP, "Completed Time": DT(2025, 6, 3, 10, 0),
           "Settlement Date": DT(2025, 6, 3)})
assert h.tanggal == D(2025, 6, 3) and h.sumber == "COMPLETED TIME" and h.selisih_abaikan, h
# blank Completed Time -> Settlement Date (even with an explicit Completed Date)
h = reb(**{"Completed Date": SEP, "Completed Time": None, "Settlement Date": SEP + 4})
assert h.tanggal == D(2026, 9, 5) and h.sumber == "SETTLEMENT DATE" and not h.selisih_abaikan, h
# bare time-of-day in Completed Time is not a date -> Settlement Date
h = reb(**{"Completed Date": None, "Completed Time": 0.4, "Settlement Date": SEP + 4})
assert h.tanggal == D(2026, 9, 5) and h.sumber == "SETTLEMENT DATE", h
# nothing usable -> no date, and Completed Date is STILL not used; Apply Date never
h = reb(**{"Completed Date": SEP + 2, "Completed Time": None, "Settlement Date": None,
           "Apply Date": SEP})
assert h.tanggal is None and h.alasan_utama == "kosong" and h.aturan == "rebate", h
# 1970 Completed Date on a Rebate row is simply ignored (Completed Time wins)
h = reb(**{"Completed Date": DT(1970, 1, 1), "Completed Time": SEP + 8.1})
assert h.tanggal == D(2026, 9, 9) and h.sumber == "COMPLETED TIME", h
# midnight crossing: applied 30 Sep, completed 00:00:20 on 1 Oct -> October by design
h = reb(**{"Completed Date": None, "Completed Time": DT(2026, 10, 1, 0, 0, 20),
           "Settlement Date": SEP + 29})
assert h.tanggal == D(2026, 10, 1) and h.sumber == "COMPLETED TIME", h
# source name: case-insensitive, trimmed, inner spaces collapsed
h = reb("  rebate   WITHDRAWAL ", **{"Completed Date": DT(2026, 8, 10), "Completed Time": SEP + 3.5})
assert h.tanggal == D(2026, 9, 4) and h.aturan == "rebate", h
# Withdrawal / other source / missing / ambiguous Source Name -> v1.28 rule (Completed Date wins)
for src in ("Withdrawal", "08-2026 Withdrawal", "Rebate Withdrawals", "Rebate", "", None, 5):
    h = reb(src, **{"Completed Date": DT(2026, 8, 10), "Completed Time": SEP + 3.5,
                    "Settlement Date": SEP + 3})
    assert h.tanggal == D(2026, 8, 10) and h.sumber == "COMPLETED DATE" and h.aturan is None, (src, h)
    assert h.konflik == D(2026, 9, 4), (src, h)                  # v1.28 conflict flag unchanged
h = hasil({"Completed Date": DT(2026, 8, 10), "Completed Time": SEP + 3.5})   # no Source Name key
assert h.tanggal == D(2026, 8, 10) and h.aturan is None, h
# Withdrawal with blank Completed Date still falls back to Completed Time (v1.28)
h = reb("Withdrawal", **{"Completed Date": None, "Completed Time": SEP + 3.5})
assert h.tanggal == D(2026, 9, 4) and h.sumber == "COMPLETED TIME" and h.aturan is None, h
# UTAMAKAN_WAKTU only affects Withdrawal; Rebate never uses Completed Date either way
tu.UTAMAKAN_WAKTU = True
try:
    h = reb("Withdrawal", **backfill)
    assert h.tanggal == D(2025, 3, 3), h
    h = reb(**backfill)
    assert h.tanggal == D(2025, 3, 3) and h.aturan == "rebate", h
finally:
    tu.UTAMAKAN_WAKTU = False
# other primary columns ignore Source Name
h = tu.tentukan_tanggal({"PAID DATE": None, "SETTLEMENT DATE": SEP, "SOURCE NAME": "Rebate Withdrawal"},
                        "PAID DATE")
assert h.tanggal is None and h.aturan is None, h
# the columns a caller must read
assert tu.kolom_penentu("COMPLETED DATE") == ("COMPLETED DATE", "COMPLETED TIME",
                                              "SETTLEMENT DATE", "SOURCE NAME")
assert tu.kolom_penentu("PAID DATE") == ("PAID DATE",) and tu.kolom_penentu(None) == ()
# Date Check counter: only rows whose Completed Date differs from Completed Time by > 3 days
ring = tu.Ringkasan("W")
ring.catat(reb(**{"Completed Date": DT(2026, 8, 10), "Completed Time": SEP + 3.5}), "COMPLETED DATE", True)
ring.catat(reb(**{"Completed Date": SEP + 3, "Completed Time": SEP + 3.5}), "COMPLETED DATE", True)   # same day
ring.catat(reb(**{"Completed Date": SEP + 6, "Completed Time": SEP + 3.5}), "COMPLETED DATE", True)   # 3 days: not counted
ring.catat(reb(**{"Completed Date": SEP + 8, "Completed Time": SEP + 3.5}), "COMPLETED DATE", True)   # 5 days: counted
ring.catat(reb(**{"Completed Date": None, "Completed Time": SEP + 3.5}), "COMPLETED DATE", True)
ring.catat(reb(**{"Completed Date": SEP, "Completed Time": None, "Settlement Date": SEP}), "COMPLETED DATE", True)
assert ring.rebate_selisih == 2 and ring.rebate_sumber == {"COMPLETED TIME": 5, "SETTLEMENT DATE": 1}, \
    (ring.rebate_selisih, ring.rebate_sumber)
assert not ring.cadangan, ring.cadangan              # Rebate is a rule, not a "fallback"
baris = ring.baris_ringkasan()
assert ("W", "Completed Date ignored for Rebate",
        "Completed Date differs from Completed Time by more than 3 days (informational)", 2) in baris, baris
assert not any(b[1] == "WARNING" for b in baris), baris

# ------------------------------------------------------------------ unit: summary / warning
ring = tu.Ringkasan("W")
for _ in range(98):
    ring.catat(hasil({"Completed Date": SEP + 1}), "COMPLETED DATE", True)
assert not ring.lewat_ambang()
for _ in range(2):
    ring.catat(hasil({"Completed Date": None}), "COMPLETED DATE", False)
assert ring.total == 100 and ring.n_tanpa == 2 and ring.lewat_ambang()          # 2% > 1%
baris = ring.baris_ringkasan()
assert any(b[1] == "WARNING" for b in baris), baris
ring2 = tu.Ringkasan("W")
for _ in range(199):
    ring2.catat(hasil({"Completed Date": SEP + 1}), "COMPLETED DATE", True)
ring2.catat(hasil({"Completed Date": None}), "COMPLETED DATE", False)
assert not ring2.lewat_ambang()                                                   # 0.5%

# ------------------------------------------------------------------ end to end (plain path)
W_HEAD = ["Source Name", "Sales UID", "Sales Name", "Sales Dept", "Sales Area", "User ID",
          "Account", "Apply Date", "Apply Time", "Currency", "Payment Gateway", "Reference",
          "USD", "Rate", "Charges", "Transaction", "Trading Charges", "Pay Channel",
          "Completed Date", "Completed Time", "Settlement Date", "Status", "Remark"]
D_HEAD = ["Source Name", "Sales UID", "User ID", "Account", "Paid Date", "Currency",
          "Payment Gateway", "Reference", "USD", "Rate", "Transaction", "Settlement Date",
          "Status"]
ref = [1000]


def w_row(src, uid, apply, cdate, ctime, settle, status="finish", pay="USDT-TRC20",
          cur="USDT", amount=100.0):
    ref[0] += 1
    return [src, uid, "S", "S", "S", ref[0], "A", apply, 0.5, cur, "USDT-TRC20", ref[0],
            amount, 1, 0, amount, 0, pay, cdate, ctime, settle, status, None]


def buat_sumber(rows_w, rows_d=None):
    wb = Workbook()
    wb.remove(wb.active)
    d = wb.create_sheet("D")
    d.append(D_HEAD)
    for r in rows_d or [["Deposit", 1, 1, "A", DT(2026, 9, 3), "USDT", "USDT-TRC20", 1, 50, 1, 50,
                         DT(2026, 9, 3), "finish"]]:
        d.append(r)
    w = wb.create_sheet("W")
    w.append(W_HEAD)
    for r in rows_w:
        w.append(r)
    x = wb.create_sheet("Xero")
    x.append(["Date", "USDT", "USD"])
    for day in range(1, 31):
        x.append([DT(2026, 9, day), 1, 1])
    f = wb.create_sheet("D&W TD FEE")
    f.append(["Dupoin Markets"])
    f.append([" Handling Fee Rate"])
    f.append(["Currency", "Payment Gateway", "Deposit", "Withdrawal"])
    f.append(["USDT", "USDT-TRC20", 0.01, 0.02])
    return wb


def jalan(*args):
    r = subprocess.run([PY, *map(str, args)], cwd=HERE, capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    return r.stdout


def pipeline(tmp, wb, bulan="2026-09", lewat_kvb=False):
    wb.save(tmp / "src.xlsx")
    sumber = tmp / "src.xlsx"
    if lewat_kvb:
        jalan("siapkan_kvb.py", sumber, "-o", tmp / "t.xlsx")
        sumber = tmp / "t.xlsx"
    jalan("buat_template.py", sumber, "-o", tmp / "tpl0.xlsx")
    o1 = jalan("isi_template.py", sumber, "-t", tmp / "tpl0.xlsx", "--bulan", bulan,
               "--tanpa-extra-fees", "-o", tmp / "tpl.xlsx")
    o2 = jalan("hitung_dw.py", tmp / "tpl.xlsx", "--period", bulan, "-o", tmp / "hasil.xlsx",
               "--tanpa-extra-fees")
    return o1, o2


def baca(path, sheet):
    wb = load_workbook(path, read_only=True, data_only=True)
    rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    wb.close()
    return rows


def detail_wd(path):
    rows = baca(path, "D&W Detail")
    h = rows[0]
    out = [dict(zip(h, r)) for r in rows[1:] if r[h.index("Type")] == "withdrawal"]
    return out


def teks_sheet(path, sheet):
    return "\n".join(" | ".join(str(c) for c in r if c is not None)
                     for r in baca(path, sheet))


def isi_w():
    return [
        # 1 regular row (valid Completed Date in Sep)
        w_row("Withdrawal", 100001, SEP + 1, DT(2026, 9, 2), DT(2026, 9, 2, 8, 0), SEP + 1),
        # 2 blank Completed Date + datetime Completed Time (Sep) -> rescued
        w_row("Rebate Withdrawal", 100007, SEP + 2, None, SEP + 4.4, SEP + 3, pay="success"),
        # 3 text 'success' Completed Date (shifted export), Status filled -> rescued
        w_row("Rebate Withdrawal", 100007, SEP + 3, "success", SEP + 5.6, SEP + 4,
              status="success", pay="success"),
        # 4 Settlement-only -> rescued
        w_row("Rebate Withdrawal", 100058, SEP + 4, None, None, SEP + 6),
        # 5 1970 Completed Date, no Completed Time -> stays excluded + reported
        w_row("Withdrawal", 100001, SEP + 5, DT(1970, 1, 1), None, SEP + 7),
        # 6 blank everything -> reported as 'no usable date'
        w_row("Rebate Withdrawal", 100016, SEP + 6, None, None, None),
        # 7 blank Completed Date, Completed Time in AUGUST -> counted as outside month (Aug)
        w_row("Rebate Withdrawal", 100007, AUG + 3, None, AUG + 4.5, AUG + 4),
        # 8 backfilled row: explicit Sep Completed Date but Completed Time in Mar 2025
        w_row("Rebate Withdrawal", 100007, SEP + 7, DT(2026, 9, 8), DT(2025, 3, 3, 10, 0),
              DT(2025, 3, 3)),
        # 9 explicit Aug Completed Date -> outside month
        w_row("Withdrawal", 100001, AUG + 1, DT(2026, 8, 2), DT(2026, 8, 2, 8, 0), AUG + 1),
        # 10 shifted row whose Status is blank but a status word sits under Pay Channel
        w_row("Rebate Withdrawal", 100007, SEP + 8, None, SEP + 9.3, SEP + 8,
              status=None, pay="success"),
    ]


tmp = Path(tempfile.mkdtemp())
o1, o2 = pipeline(tmp, buat_sumber(isi_w()))
o1, o2 = pipeline(tmp, buat_sumber(isi_w()))
# v1.29: rows 2,3,4,10 + the regular row are in; row 8 (Rebate, Completed Date 8 Sep but
# Completed Time Mar 2025) is no longer pulled into September by its fake Completed Date.
assert "sheet W: 5 baris dari 'W', 5 di luar 2026-09 dilewati" in o1, o1
assert "dilewati karena di luar 2026-09: 3 baris (2025-03: 1, 2026-08: 2)" in o1, o1
assert "Rebate Withdrawal: 'Completed Date' DIABAIKAN, tanggal dari Completed Time: 5, " \
       "Settlement Date: 1 (1 baris" in o1, o1
assert "kolom cadangan" not in o1, o1                       # no Withdrawal row needed a fallback
assert "TANPA tanggal yang bisa dipakai: 2 baris" in o1, o1
assert "2 dari 10 baris (20.0%)" in o1 and "PERINGATAN" in o1, o1             # > 1% warning (log)
assert "3 baris punya kata status" in o1, o1                                   # shifted columns
assert "(1970" in o1 or "1970" in o1, o1

# the rows reach the worksheet 'W' of the template and the report
isi = baca(tmp / "tpl.xlsx", "W")
assert len(isi) - 1 == 5, len(isi) - 1
dtl = detail_wd(tmp / "hasil.xlsx")
assert len(dtl) == 5, dtl
por = sorted((r["Sales UID"], r["Date Source"]) for r in dtl)
assert por == sorted([(100001, "Completed Date"), (100007, "Completed Time"),
                      (100007, "Completed Time"), (100058, "Settlement Date"),
                      (100007, "Completed Time")]), por
assert {r["Date"].month for r in dtl} == {9}
# backfilled Rebate row (row 8): its fake Sep Completed Date is IGNORED -> not in September
assert not any(r["Date"] == DT(2026, 9, 8) for r in dtl), dtl
# shifted row with blank Status takes the status word from Pay Channel -> counted, not refused
assert sum(1 for r in dtl if str(r["Status"]).lower() in ("success", "finish")) == 5
# totals: 5 withdrawals x 100 USD
rpt = teks_sheet(tmp / "hasil.xlsx", "D&W Report")
assert "500" in rpt, rpt

# Missing Data + Date Check sheets say what was skipped, by month and reason
md = teks_sheet(tmp / "hasil.xlsx", "Missing Data")
assert "5 rows were already skipped" in md, md
assert "skipped: outside report month: 2026-08" in md, md
assert "skipped: outside report month: 2025-03" in md, md
assert "skipped: no usable date: blank" in md and "outside 1990-2100" in md, md
assert "Rebate Withdrawal: Completed Date ignored, date taken from: Completed Time" in md, md
assert "Completed Date ignored for Rebate" in md, md
assert "WARNING: 20.0% of rows have no usable date (limit 1%)" in md, md
assert "shifted columns: status word found under 'Pay Channel'" in md, md
assert "None - every row in the file already fell inside this month" not in md
dc = teks_sheet(tmp / "hasil.xlsx", "Date Check")
assert "skipped: outside report month | 2026-08 | 2" in dc, dc
assert "Completed Date ignored for Rebate | Completed Date differs from Completed Time by more " \
       "than 3 days (informational) | 1" in dc, dc
ver = baca(tmp / "hasil.xlsx", "Verifikasi Tanggal")
assert any(r[0] == "W" and r[1] == 6 for r in ver[2:]), ver                  # 1970 row listed

# a clean month: no rows lost -> no warning, Missing Data stays quiet
bersih = [w_row("Withdrawal", 100001, SEP + i, DT(2026, 9, i + 1), None, SEP + i)
          for i in range(5)]
tmp2 = Path(tempfile.mkdtemp())
o1b, o2b = pipeline(tmp2, buat_sumber(bersih))
assert "TANPA tanggal" not in o1b and "punya kata status" not in o1b and "BEDA" not in o1b, o1b
assert "kolom cadangan" not in o1b, o1b
md2 = teks_sheet(tmp2 / "hasil.xlsx", "Missing Data")
assert "WARNING" not in md2 and "skipped" not in md2, md2
assert len(detail_wd(tmp2 / "hasil.xlsx")) == 5

# exactly 1% (1 of 100) is not a warning; 2 of 100 is
tmp3 = Path(tempfile.mkdtemp())
cien = [w_row("Withdrawal", 100001, SEP + 1, DT(2026, 9, 2), None, SEP + 1) for _ in range(99)]
cien.append(w_row("Rebate Withdrawal", 100016, SEP + 1, None, None, None))
o1c, _ = pipeline(tmp3, buat_sumber(cien))
assert "TANPA tanggal yang bisa dipakai: 1 baris" in o1c and "20" not in o1c.split("TANPA")[1][:60]
assert "(1.0%)" not in o1c, o1c

# ------------------------------------------------------------------ KVB path
# siapkan_kvb -> isi_template -> hitung_dw is the SAME code, so the same rows come back.
# A raw FLOAT serial in 'Completed Time' (siapkan_kvb only converts headers with DATE).
tmp4 = Path(tempfile.mkdtemp())
src = Workbook()
src.remove(src.active)
t = src.create_sheet("Transfer")
t.append(["汇出"] + [None] * 7 + ["收款"])
t.append(["汇款日期", "汇出部门", "支付渠道", "汇出币种", "Xero Rate", "金额", "USD (Xero)", "手续费",
          "收款日期", "收款部门", "收款人", "收入币种", "金额", "手续费", "USD (Xero)"])
t.append([AUG + 3, "KVB", "ChipPay", "CNY", None, 1000, None, None,
          AUG + 3, "KVB", "J Wallet", "USDT", 140, None, None])
d = src.create_sheet("D")
d.append(["Source.Name", "user_id", "account", "Paid Date", "Settlement Date", "method", "ticket",
          "USD (CRM)", "rate", "transaction", "receive_company_account"])
d.append(["08-2026 Deposit", 1, 1, DT(2026, 8, 3), DT(2026, 8, 3), "CNY", 1, 140, 0.14, 1000,
          "ChipPay"])
w = src.create_sheet("W")
w.append(["Source.Name", "User ID", "Account", "Apply Date", "Payment Currency", "Payment Gateway",
          "Reference", "USD (CRM)", "rate", "charges", "transaction", "Completed Date",
          "Completed Time", "Settlement Date", "status"])
w.append(["08-2026 Withdrawal", 1, 1, AUG + 5, "CNY", "ChipPay", 201, 140, 0.14, 0, 1000,
          AUG + 5, None, AUG + 5, "finish"])                        # normal
w.append(["08-2026 Withdrawal", 2, 2, AUG + 6, "CNY", "ChipPay", 202, 70, 0.14, 0, 500,
          None, AUG + 7.55, AUG + 6, "finish"])                     # blank date, float Completed Time
w.append(["08-2026 Withdrawal", 3, 3, AUG + 7, "CNY", "ChipPay", 203, 70, 0.14, 0, 500,
          "success", DT(2026, 8, 9, 10, 30), AUG + 8, "finish"])    # text date, datetime time
w.append(["08-2026 Withdrawal", 4, 4, AUG + 8, "CNY", "ChipPay", 204, 14, 0.14, 0, 100,
          None, None, AUG + 9, "finish"])                           # Settlement only
w.append(["08-2026 Withdrawal", 5, 5, AUG + 9, "CNY", "ChipPay", 205, 14, 0.14, 0, 100,
          None, None, None, "finish"])                              # nothing usable
w.append(["07-2026 Withdrawal", 6, 6, AUG - 5, "CNY", "ChipPay", 206, 14, 0.14, 0, 100,
          None, AUG - 4.5, AUG - 4, "finish"])                      # July via Completed Time
w.append(["Rebate Withdrawal", 7, 7, AUG + 3, "CNY", "ChipPay", 207, 14, 0.14, 0, 100,
          AUG + 4, AUG - 3.5, AUG + 3, "finish"])                   # fake Aug Completed Date, real July
w.append(["Rebate Withdrawal", 8, 8, AUG + 9, "CNY", "ChipPay", 208, 14, 0.14, 0, 100,
          None, AUG + 10.3, AUG + 9, "finish"])                     # Rebate, blank date -> Completed Time
x = src.create_sheet("Xero")
x.append(["Date", "CNY", "USDT", "USD"])
for day in range(1, 32):
    x.append([DT(2026, 8, day), 7.0, 1.0, 1.0])
f = src.create_sheet("Handing Fee")
f.append(["Payment Channel", "Currency", "Deposit", "Withdrawal"])
f.append(["ChipPay", "CNY", 0.004, 0.007])
src.save(tmp4 / "kvb.xlsx")

jalan("siapkan_kvb.py", tmp4 / "kvb.xlsx", "-o", tmp4 / "t.xlsx")
jalan("buat_template.py", tmp4 / "t.xlsx", "-o", tmp4 / "tpl0.xlsx")
ok = jalan("isi_template.py", tmp4 / "t.xlsx", "-t", tmp4 / "tpl0.xlsx", "--bulan", "2026-08",
           "--tanpa-extra-fees", "-o", tmp4 / "tpl.xlsx")
assert "sheet W: 5 baris dari 'W', 3 di luar 2026-08 dilewati" in ok, ok
assert "Completed Time: 3, Settlement Date: 1" in ok, ok            # the Withdrawal fallbacks
assert "Rebate Withdrawal: 'Completed Date' DIABAIKAN, tanggal dari Completed Time: 2 " \
       "(1 baris" in ok, ok
assert "TANPA tanggal yang bisa dipakai: 1 baris" in ok, ok
jalan("hitung_dw.py", tmp4 / "tpl.xlsx", "--period", "2026-08", "-o", tmp4 / "hasil.xlsx",
      "--tanpa-extra-fees")
kd = detail_wd(tmp4 / "hasil.xlsx")
assert sorted(r["Reference"] for r in kd) == [201, 202, 203, 204, 208], kd
assert {r["Date Source"] for r in kd} == {"Completed Date", "Completed Time", "Settlement Date"}
assert all(r["Date"].month == 8 for r in kd)
kmd = teks_sheet(tmp4 / "hasil.xlsx", "Missing Data")
assert "skipped: outside report month: 2026-07" in kmd, kmd
assert "skipped: no usable date: blank" in kmd, kmd
assert "WARNING: 12.5% of rows have no usable date" in kmd, kmd
assert "Completed Date ignored for Rebate" in kmd, kmd
assert 207 not in {r["Reference"] for r in kd}                   # fake Aug Completed Date ignored on KVB too

# ------------------------------------------------------------------ v1.29 rule, end to end
# Each row is one case; September is the report month.
regel = [
    # A Rebate: Completed Date Aug, Completed Time Sep -> Sep (Completed Time)
    w_row("Rebate Withdrawal", 100007, SEP + 1, DT(2026, 8, 10), SEP + 3.5, SEP + 3),
    # B Rebate: Completed Date 1 Sep, Completed Time in 2025 -> NOT September
    w_row("Rebate Withdrawal", 100007, SEP + 1, SEP, DT(2025, 6, 3, 10, 0), DT(2025, 6, 3)),
    # C Rebate: Completed Time blank -> Settlement Date (5 Sep)
    w_row("Rebate Withdrawal", 100007, SEP + 2, SEP, None, SEP + 4),
    # D Rebate: nothing usable except a fake Completed Date -> reported, not counted
    w_row("Rebate Withdrawal", 100007, SEP + 2, SEP + 2, None, None),
    # E Withdrawal: Completed Date 30 Aug wins over Completed Time 2 Sep -> August, outside
    w_row("Withdrawal", 100001, SEP, DT(2026, 8, 30), SEP + 1.4, SEP + 1),
    # F Withdrawal: blank Completed Date -> Completed Time (v1.28)
    w_row("Withdrawal", 100001, SEP + 4, None, SEP + 5.4, SEP + 5),
    # G Rebate, midnight crossing: applied 30 Sep, completed 00:00:20 on 1 Oct -> October
    w_row("Rebate Withdrawal", 100007, SEP + 29, None, DT(2026, 10, 1, 0, 0, 20), SEP + 29),
    # H Rebate with sloppy Source Name spelling -> still Rebate
    w_row("  rebate WITHDRAWAL ", 100007, SEP + 6, DT(2026, 8, 12), SEP + 7.2, SEP + 7),
    # I Source Name missing -> v1.28 rule: Completed Date (Aug) wins -> outside
    w_row(None, 100007, SEP + 6, DT(2026, 8, 12), SEP + 7.2, SEP + 7),
    # J ambiguous Source Name -> v1.28 rule
    w_row("Rebate Withdrawals", 100007, SEP + 6, DT(2026, 8, 12), SEP + 7.2, SEP + 7),
]
tmp6 = Path(tempfile.mkdtemp())
o1f, o2f = pipeline(tmp6, buat_sumber(regel))
assert "sheet W: 4 baris dari 'W', 6 di luar 2026-09 dilewati" in o1f, o1f
assert "Rebate Withdrawal: 'Completed Date' DIABAIKAN, tanggal dari Completed Time: 4, " \
       "Settlement Date: 1 (3 baris" in o1f, o1f                  # A, B, H differ by > 3 days
assert "Completed Time: 1" in o1f and "kolom cadangan" in o1f, o1f   # Withdrawal row F
fd = detail_wd(tmp6 / "hasil.xlsx")
assert sorted((r["Date"].day, r["Date Source"]) for r in fd) == sorted(
    [(4, "Completed Time"), (5, "Settlement Date"), (6, "Completed Time"), (8, "Completed Time")]), fd
assert all(r["Date"].month == 9 for r in fd)
# the template copy and the report agree: same 4 rows reach tab W and the report
assert len(baca(tmp6 / "tpl.xlsx", "W")) - 1 == 4
fmd = teks_sheet(tmp6 / "hasil.xlsx", "Missing Data")
assert "6 rows were already skipped" in fmd, fmd
dcf = teks_sheet(tmp6 / "hasil.xlsx", "Date Check")
assert "Completed Date ignored for Rebate | Completed Date differs from Completed Time by more " \
       "than 3 days (informational) | 3" in dcf, dcf
assert "skipped: outside report month | 2025-06 | 1" in dcf, dcf
assert "skipped: outside report month | 2026-10 | 1" in dcf, dcf          # midnight crossing
assert "skipped: outside report month | 2026-08 | 3" in dcf, dcf          # E, I, J
assert "skipped: no usable date | blank | 1" in dcf, dcf                   # D
# MTOATD uses the SAME date as the month filter (not the raw Completed Date of Rebate rows):
# date blocks = 3 Sep (the deposit) + the four dates above; raw Completed Dates (1, 2 Sep,
# 10 Aug ...) must not create blocks.
mt = baca(tmp6 / "hasil.xlsx", "MTOATD")
blok = [r[1].day for r in mt if r[0] == "日期：" and isinstance(r[1], datetime.datetime)]
assert blok == [3, 4, 5, 6, 8], blok
# 'Completed Date' of D is the fake one (2 Sep) but D is NOT counted -> W has no 100007 row on 3 Sep
assert not any(r["Reference"] == regel[3][11] for r in fd)

# D (Paid Date) is unchanged: a blank Paid Date does NOT borrow Settlement Date
tmp5 = Path(tempfile.mkdtemp())
dd = [["Deposit", 1, 1, "A", None, "USDT", "USDT-TRC20", 1, 50, 1, 50, DT(2026, 9, 3), "finish"],
      ["Deposit", 2, 2, "A", DT(2026, 9, 4), "USDT", "USDT-TRC20", 2, 50, 1, 50, DT(2026, 9, 4),
       "finish"]]
o1e, _ = pipeline(tmp5, buat_sumber([w_row("Withdrawal", 100001, SEP, DT(2026, 9, 2), None, SEP)], dd))
assert "sheet D: 1 baris dari 'D'" in o1e, o1e

print("Completed-date fallback (Completed Time -> Settlement), 1970, backfill, Missing Data, "
      "shifted status, >1% warning, KVB path, v1.29 Rebate rule (Completed Date ignored): OK")
