"""v1.44 (DPM): Withdrawal AND Rebate/Wallet Withdrawal are dated by Completed Date;
withdrawal Status finish / success / process is counted (deposits: finish / success);
'Wallet Withdrawal' (new DPM name from Nov 2026) == 'Rebate Withdrawal' everywhere.
KVB keeps the old rules through --legacy-withdrawal-rules (app.py KVB_ARG).

Synthetic data only (no client data).    python3 test_dpm_withdrawal_completed_date_v144.py
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


def muat(nama):
    spec = importlib.util.spec_from_file_location(nama + "_v144", HERE / f"{nama}.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


hd = muat("hitung_dw")
spec = muat("mtoatd_spec")
app_src = (HERE / "app.py").read_text()

assert tu.REBATE_PAKAI_COMPLETED_DATE is True                 # DPM default


def hasil(src, **kw):
    n = {"SOURCE NAME": src}
    n.update({k.upper(): v for k, v in kw.items()})
    return tu.tentukan_tanggal(n, "COMPLETED DATE")


REBATES = ("Rebate Withdrawal", "Wallet Withdrawal", "  wallet   WITHDRAWAL ", "08-2026 Wallet withdrawal")

# ------------------------------------------------------------------ unit: date rule (DPM)
for src in REBATES:
    assert tu.adalah_rebate(src), src
    # time-only Completed Time (text, as in the Oct 2026 export), Completed Date = Settlement + 1
    h = hasil(src, **{"Completed Date": DT(2026, 10, 3), "Completed Time": " 00:33",
                      "Settlement Date": DT(2026, 10, 2)})
    assert h.tanggal == D(2026, 10, 3) and h.sumber == "COMPLETED DATE" and h.aturan is None, (src, h)
    # same with a bare time value
    h = hasil(src, **{"Completed Date": DT(2026, 10, 11), "Completed Time": datetime.time(9, 38),
                      "Settlement Date": DT(2026, 10, 10)})
    assert h.tanggal == D(2026, 10, 11) and h.sumber == "COMPLETED DATE", (src, h)
    # full-date Completed Time within 3 days (midnight crossing): Completed Date kept
    h = hasil(src, **{"Completed Date": DT(2026, 10, 1), "Completed Time": DT(2026, 9, 30, 23, 59),
                      "Settlement Date": DT(2026, 9, 30)})
    assert h.tanggal == D(2026, 10, 1) and h.sumber == "COMPLETED DATE" and not h.selisih_abaikan, h
    # v1.29 back-filled Completed Date (far from a full-date Completed Time) -> Completed Time
    h = hasil(src, **{"Completed Date": DT(2026, 8, 10), "Completed Time": DT(2025, 11, 4, 10, 0),
                      "Settlement Date": DT(2025, 11, 4)})
    assert h.tanggal == D(2025, 11, 4) and h.sumber == "COMPLETED TIME" and h.aturan == "rebate" \
        and h.selisih_abaikan, (src, h)
    # blank Completed Date: full Completed Time -> it; time-only -> Settlement Date (fallback)
    h = hasil(src, **{"Completed Date": None, "Completed Time": DT(2026, 10, 4, 1, 0),
                      "Settlement Date": DT(2026, 10, 3)})
    assert h.tanggal == D(2026, 10, 4) and h.sumber == "COMPLETED TIME" and h.aturan is None, h
    h = hasil(src, **{"Completed Date": None, "Completed Time": " 07:00",
                      "Settlement Date": DT(2026, 10, 3)})
    assert h.tanggal == D(2026, 10, 3) and h.sumber == "SETTLEMENT DATE", h
    # 1970 = not paid yet: never rescued by Settlement Date (same as regular Withdrawal)
    h = hasil(src, **{"Completed Date": DT(1970, 1, 1), "Completed Time": " 07:00",
                      "Settlement Date": DT(2026, 10, 31)})
    assert h.tanggal is None and h.alasan_utama == "tidak_sah", h
    h = hasil(src, **{"Completed Date": DT(1970, 1, 1), "Completed Time": DT(2026, 10, 9, 8, 0),
                      "Settlement Date": DT(2026, 10, 8)})
    assert h.tanggal == D(2026, 10, 9) and h.sumber == "COMPLETED TIME", h
# 'Wallet Withdrawal' and 'Rebate Withdrawal' give identical results for the same values
for kw in ({"Completed Date": DT(2026, 10, 3), "Completed Time": " 00:33", "Settlement Date": DT(2026, 10, 2)},
           {"Completed Date": DT(2026, 8, 10), "Completed Time": DT(2025, 11, 4), "Settlement Date": None},
           {"Completed Date": None, "Completed Time": None, "Settlement Date": DT(2026, 10, 2)}):
    assert hasil("Rebate Withdrawal", **kw) == hasil("Wallet Withdrawal", **kw), kw
# regular Withdrawal: Completed Date wins over Settlement Date (day before)
h = hasil("Withdrawal", **{"Completed Date": DT(2026, 10, 2), "Completed Time": " 00:10",
                           "Settlement Date": DT(2026, 10, 1)})
assert h.tanggal == D(2026, 10, 2) and h.sumber == "COMPLETED DATE", h
for src in ("Wallet Withdrawals", "Wallet", "Withdrawal", None):
    assert not tu.adalah_rebate(src) or src is None, src

# Date Check: only the implausible rebate dates are listed, with the new label
ring = tu.Ringkasan("W")
ring.catat(hasil("Wallet Withdrawal", **{"Completed Date": DT(2026, 8, 10),
                                         "Completed Time": DT(2026, 10, 4, 1)}), "COMPLETED DATE", True)
ring.catat(hasil("Rebate Withdrawal", **{"Completed Date": DT(2026, 10, 3),
                                         "Completed Time": " 00:33"}), "COMPLETED DATE", True)
ring.catat(hasil("Rebate Withdrawal", **{"Completed Date": None, "Completed Time": " 00:33",
                                         "Settlement Date": DT(2026, 10, 2)}), "COMPLETED DATE", True)
assert ring.rebate_sumber == {"COMPLETED TIME": 1} and ring.cadangan == {"SETTLEMENT DATE": 1}, \
    (ring.rebate_sumber, ring.cadangan)
baris = ring.baris_ringkasan()
assert ("W", "Rebate/Wallet Withdrawal: implausible Completed Date replaced",
        "Completed Date differs from Completed Time by more than 3 days -> date taken from "
        "Completed Time", 1) in baris, baris
assert ("W", "date taken from fallback column", "Settlement Date", 1) in baris, baris
assert not any("Completed Date ignored" in b[1] for b in baris), baris

# legacy (KVB) rule is still available and unchanged
tu.REBATE_PAKAI_COMPLETED_DATE = False
try:
    h = hasil("Rebate Withdrawal", **{"Completed Date": DT(2026, 10, 3), "Completed Time": " 00:33",
                                      "Settlement Date": DT(2026, 10, 2)})
    assert h.tanggal == D(2026, 10, 2) and h.sumber == "SETTLEMENT DATE" and h.aturan == "rebate", h
finally:
    tu.REBATE_PAKAI_COMPLETED_DATE = True

# ------------------------------------------------------------------ unit: status filter
assert hd.status_dipakai("withdrawal") == ("FINISH", "SUCCESS", "PROCESS")
assert hd.status_dipakai("deposit") == ("FINISH", "SUCCESS")
for st in ("process", " Process ", "PROCESS", "finish", "success"):
    assert hd.norm(st) in hd.status_dipakai("withdrawal"), st
for st in ("refuse", "", None, "processing", "cancel"):
    assert hd.norm(st) not in hd.status_dipakai("withdrawal"), st
assert hd.kind_of("W", "Wallet Withdrawal") == "withdrawal" == hd.kind_of("W", "Rebate Withdrawal")

# ------------------------------------------------------------------ unit: MTOATD Wallet出金
idx = spec.Indeks()
for src in ("Rebate Withdrawal", "Wallet Withdrawal", " wallet withdrawal ", "Withdrawal"):
    idx.tambah_withdrawal("USDT", D(2026, 10, 3), D(2026, 10, 2), 10.0, 10.0, 0, "success", src)
v = spec.hitung(idx, "USDT", D(2026, 10, 2))
assert v["wallet_w"] == 30.0 and v["mt4_w"] == 10.0, v

# ------------------------------------------------------------------ KVB keeps the old rules
kvb_arg = app_src[app_src.index("KVB_ARG = {"):]
kvb_arg = kvb_arg[:kvb_arg.index("\n}\n")]
assert kvb_arg.count('"--legacy-withdrawal-rules"') == 2, kvb_arg      # isi_template + hitung_dw

# ------------------------------------------------------------------ end to end (October)
W_HEAD = ["Source Name", "Sales UID", "Sales Name", "Sales Dept", "Sales Area", "User ID",
          "Account", "Apply Date", "Apply Time", "Currency", "Payment Gateway", "Reference",
          "USD", "Rate", "Charges", "Transaction", "Trading Charges", "Pay Channel",
          "Completed Date", "Completed Time", "Settlement Date", "Status", "Remark"]
D_HEAD = ["Source Name", "Sales UID", "User ID", "Account", "Paid Date", "Currency",
          "Payment Gateway", "Reference", "USD", "Rate", "Transaction", "Settlement Date",
          "Status"]
O = lambda d: DT(2026, 10, d)                      # noqa: E731
RB = "Rebate Withdrawal"
# (source, ref, transaction, completed date, completed time, settlement, status)
BARIS_W = [
    ("Withdrawal", 1, 1000.0, O(2), " 00:10", O(1), "finish"),     # regular: Completed Date 2 Oct
    (RB, 2, 2000.0, O(3), " 00:33", O(2), "success"),               # Bayu's case -> 3 Oct (was 2 Oct)
    (RB, 3, 3000.0, O(2), " 23:04", O(2), "success"),               # same day
    (RB, 4, 400.0, O(5), " 10:00", O(5), "process"),                # NEW: process counted
    ("Withdrawal", 5, 500.0, O(5), " 11:00", O(5), " Process "),    # NEW: trimmed / case
    ("Withdrawal", 6, 600.0, O(5), " 12:00", O(5), "refuse"),       # never counted
    (RB, 7, 700.0, DT(1970, 1, 1), " 07:00", O(31), "process"),     # 1970 -> no usable date
    (RB, 8, 800.0, DT(2026, 11, 1), " 00:20", O(31), "success"),    # month mover -> November
    ("Withdrawal", 9, 900.0, None, " 09:00", O(6), "finish"),        # blank -> Settlement fallback
]


def buat(nama_rebate):
    wb = Workbook()
    wb.remove(wb.active)
    d = wb.create_sheet("D")
    d.append(D_HEAD)
    d.append(["Deposit", 1, 1, "A", O(2), "LAK", "THX-LAK", 91, 5.0, 20000, 100000.0, O(2), "finish"])
    d.append(["Deposit", 1, 1, "A", O(2), "LAK", "THX-LAK", 92, 5.0, 20000, 7777.0, O(2), "process"])
    w = wb.create_sheet("W")
    w.append(W_HEAD)
    for src, ref, tx, cd, ct, sd, st in BARIS_W:
        src = nama_rebate if src == RB else src
        w.append([src, 100001, "S", "S", "S", ref, "A", sd, " 00:00", "LAK", "THX-LAK", ref,
                  tx / 20000, 20000, 0, tx, 0, "LAK", cd, ct, sd, st, None])
    x = wb.create_sheet("Xero")
    x.append(["Date", "LAK", "USD"])
    for day in range(1, 32):
        x.append([O(day), 20000 + day, 1])                   # rate shows which day was used
    for day in range(1, 3):
        x.append([DT(2026, 11, day), 21000 + day, 1])
    f = wb.create_sheet("D&W TD FEE")
    f.append(["Dupoin Markets"])
    f.append([" Handling Fee Rate"])
    f.append(["Currency", "Payment Gateway", "Deposit", "Withdrawal"])
    f.append(["LAK", "THX-LAK", 0.03, 0.01])
    return wb


def jalan(*args):
    r = subprocess.run([PY, *map(str, args)], cwd=HERE, capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    return r.stdout


def pipeline(wb, *extra):
    tmp = Path(tempfile.mkdtemp())
    wb.save(tmp / "src.xlsx")
    jalan("buat_template.py", tmp / "src.xlsx", "-o", tmp / "tpl0.xlsx")
    o1 = jalan("isi_template.py", tmp / "src.xlsx", "-t", tmp / "tpl0.xlsx", "--bulan", "2026-10",
               "-o", tmp / "tpl.xlsx", *extra)
    o2 = jalan("hitung_dw.py", tmp / "tpl.xlsx", "--period", "2026-10", "-o", tmp / "hasil.xlsx",
               *extra)
    return tmp, o1, o2


def baca(path, sheet):
    wb = load_workbook(path, read_only=True, data_only=True)
    rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    wb.close()
    return rows


def detail_wd(path):
    rows = baca(path, "D&W Detail")
    h = rows[0]
    return h, [dict(zip(h, r)) for r in rows[1:] if r[h.index("Type")] == "withdrawal"]


def pcb_withdrawal(path):
    rows = baca(path, "Payment Channel Balance")
    h = rows[0]
    out = {}
    for r in rows[1:]:
        if r[h.index("Currency")] == "LAK" and isinstance(r[0], datetime.datetime) \
                and r[h.index("Payment Gateway")] == "THX":
            out[r[0].date()] = (r[h.index("Withdrawal")] or 0, r[h.index("Withdrawal charges")] or 0)
    return out


def tanggal_detail(det):
    return sorted((int(r["Reference"]), r["Date"].date() if hasattr(r["Date"], "date") else r["Date"])
                  for r in det)


tmp_r, o1_r, o2_r = pipeline(buat("Rebate Withdrawal"))
h_r, det_r = detail_wd(tmp_r / "hasil.xlsx")
assert tanggal_detail(det_r) == [(1, D(2026, 10, 2)), (2, D(2026, 10, 3)), (3, D(2026, 10, 2)),
                                 (4, D(2026, 10, 5)), (5, D(2026, 10, 5)), (9, D(2026, 10, 6))], \
    tanggal_detail(det_r)
# Xero rate is looked up on the same (Completed) date
for r in det_r:
    assert r["Xero Rate"] == 20000 + r["Date"].day, r
pcb = pcb_withdrawal(tmp_r / "hasil.xlsx")
assert pcb[D(2026, 10, 2)][0] == 4000.0 and pcb[D(2026, 10, 3)][0] == 2000.0, pcb
assert pcb[D(2026, 10, 5)][0] == 900.0 and pcb[D(2026, 10, 6)][0] == 900.0, pcb
assert abs(pcb[D(2026, 10, 2)][1] - 40.0) < 1e-9, pcb                     # 1% charge follows
# deposits are untouched by v1.44: the template's D tab has no Status column (so the D
# side is never status-filtered in this pipeline) -> both deposits, same as before
dr = baca(tmp_r / "hasil.xlsx", "D&W Detail")
dep = [dict(zip(dr[0], r)) for r in dr[1:] if r[dr[0].index("Type")] == "deposit"]
assert sorted(int(r["Reference"]) for r in dep) == [91, 92], dep
# template copy (isi_template) and report agree on the month: refs 1-6, 9 copied (status is
# filtered later in hitung_dw), 1970 and the 1 Nov row are not
tw = baca(tmp_r / "tpl.xlsx", "W")
assert sorted(int(r[tw[0].index("Reference")]) for r in tw[1:]) == [1, 2, 3, 4, 5, 6, 9], tw
dc = "\n".join(" | ".join("" if c is None else str(c) for c in r)
               for r in baca(tmp_r / "hasil.xlsx", "Date Check"))
assert "skipped: outside report month | 2026-11 | 1" in dc, dc
assert "skipped: no usable date | outside 1990-2100 (e.g. 1970) | 1" in dc, dc
assert "date taken from fallback column | Settlement Date | 1" in dc, dc
assert "Completed Date ignored" not in dc, dc
leg = "\n".join(str(c) for r in baca(tmp_r / "hasil.xlsx", "Legend") for c in r if c)
assert "'finish', 'success' or 'process'" in leg and "Wallet Withdrawal" in leg, leg
mt_r = baca(tmp_r / "hasil.xlsx", "MTOATD")

# 'Wallet Withdrawal' gives exactly the same numbers as 'Rebate Withdrawal'
tmp_w, o1_w, o2_w = pipeline(buat("Wallet Withdrawal"))
h_w, det_w = detail_wd(tmp_w / "hasil.xlsx")
assert tanggal_detail(det_w) == tanggal_detail(det_r)
tanpa_sumber = lambda det: sorted(tuple(v for k, v in r.items() if k != "Source Name")  # noqa: E731
                                  for r in det)
assert tanpa_sumber(det_w) == tanpa_sumber(det_r)
for sheet in ("Payment Channel Balance", "MTOATD", "D&W Report", "J Wallet", "Date Check"):
    assert baca(tmp_w / "hasil.xlsx", sheet) == baca(tmp_r / "hasil.xlsx", sheet), sheet
# ... and MTOATD Wallet出金 actually picks those rows up (USD on Settlement Date 2 Oct: refs 2, 3)
lbl = next(r for r in mt_r if r and any(isinstance(c, str) and c.startswith("Wallet出金") for c in r))
assert any(isinstance(c, (int, float)) and abs(c - 0.25) < 1e-9 for c in lbl), lbl

# KVB path (--legacy-withdrawal-rules): old behaviour, byte-for-byte rules of v1.43
tmp_k, o1_k, o2_k = pipeline(buat("Rebate Withdrawal"), "--legacy-withdrawal-rules")
_, det_k = detail_wd(tmp_k / "hasil.xlsx")
assert tanggal_detail(det_k) == [(1, D(2026, 10, 2)), (2, D(2026, 10, 2)), (3, D(2026, 10, 2)),
                                 (8, D(2026, 10, 31)), (9, D(2026, 10, 6))], tanggal_detail(det_k)
dk = baca(tmp_k / "hasil.xlsx", "D&W Detail")
assert sorted(int(r[dk[0].index("Reference")]) for r in dk[1:]
              if r[dk[0].index("Type")] == "deposit") == [91, 92]

print("v1.44 DPM: Withdrawal + Rebate/Wallet Withdrawal on Completed Date, status "
      "finish/success/process, Wallet == Rebate, KVB legacy unchanged: OK")

# hitung_dw run straight on a workbook whose D tab HAS a Status column: deposit 'process'
# is still dropped (deposit filter unchanged), withdrawal 'process' is counted
tmp_d = Path(tempfile.mkdtemp())
buat("Wallet Withdrawal").save(tmp_d / "src.xlsx")
jalan("hitung_dw.py", tmp_d / "src.xlsx", "--period", "2026-10", "-o", tmp_d / "hasil.xlsx")
dd = baca(tmp_d / "hasil.xlsx", "D&W Detail")
ix, it_ = dd[0].index("Reference"), dd[0].index("Type")
assert sorted(int(r[ix]) for r in dd[1:] if r[it_] == "deposit") == [91], dd
assert sorted(int(r[ix]) for r in dd[1:] if r[it_] == "withdrawal") == [1, 2, 3, 4, 5, 9], dd
print("direct hitung_dw: deposit 'process' dropped, withdrawal 'process' counted: OK")
