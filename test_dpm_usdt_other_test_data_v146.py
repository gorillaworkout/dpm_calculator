"""v1.46 (DPM): two decisions of Bayu + the DPM team.

1. A USDT row whose Payment Gateway is 'Other' is its own channel on Payment Channel
   Balance. It used to share the BEP20 row, because the fee-lookup alias
   (USDT-BEP20 / BEP20 -> OTHER) also built the channel key.
2. Back-office test rows (Sales Name 'Test data' or Sales Dept 'Test Department') are
   always excluded from Deposit, Withdrawal and Rebate / Wallet Withdrawal, and listed
   in Missing Data section 0c. KVB keeps them through --keep-test-rows (app.py KVB_ARG).

Synthetic data only (no client data).    python3 test_dpm_usdt_other_test_data_v146.py
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
DT = datetime.datetime


def muat(nama):
    spec = importlib.util.spec_from_file_location(nama + "_v146", HERE / f"{nama}.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


hd = muat("hitung_dw")
app_src = (HERE / "app.py").read_text()

# ------------------------------------------------------------------ unit: channel key
k = hd.kunci_channel
assert k("USDT", "Other") != k("USDT", "BEP20"), (k("USDT", "Other"), k("USDT", "BEP20"))
assert k("USDT", "Other") != k("USDT", "USDT-BEP20")
assert k("USDT", "USDT-BEP20") == k("USDT", "BEP20")          # BEP20 spellings still one row
assert k("USDT", " other ") == k("USDT", "OTHER") == k("USDT", "Others")   # PCB column 'OTHERS'
assert k("USDT", "Other") not in (k("USDT", "TRC20"), k("USDT", "ERC20"))
# other aliases unchanged
assert k("VND", "VN77PAY") == k("VND", "77Pay")
assert k("USD", "MNTX-BinancePay") == k("USD", "Binance Pay")
# a non-USDT 'Other' gateway keeps its old key
assert k("INR", "Other") == "OTHER"
# the fee lookup still uses the alias: USDT-BEP20 / BEP20 take the USDT/Other fee line
tabel = {("USDT", "OTHER"): {"deposit": {"pct": 0.0, "fixed": 0.0, "min": None},
                            "withdrawal": {"pct": 0.0, "fixed": 0.0, "min": None},
                            "fixed": 0.0, "nama": "Other"}}
for gw, how in (("USDT-BEP20", "alias"), ("BEP20", "alias"), ("Other", "exact")):
    komp, _f, key, cara = hd.lookup_fee(tabel, "USDT", gw, "withdrawal")
    assert komp is not None and cara == how and key == ("USDT", "Other"), (gw, cara, key)

# ------------------------------------------------------------------ unit: test-row match
assert hd.baris_data_tes("Test data", None)
assert hd.baris_data_tes("  TEST   DATA ", "Dupoin Markets Sales")
assert hd.baris_data_tes("Some Sales", "Test Department")
assert hd.baris_data_tes(None, " test department")
for nama, dept in (("Dupoin Sales", "Dupoin Markets Sales"), ("Testing Sales", "Sales"),
                   ("Test", "Test"), (None, None), ("Test data 2", "Dept")):
    assert not hd.baris_data_tes(nama, dept), (nama, dept)
assert hd.BUANG_DATA_TES is True                              # DPM default

# ------------------------------------------------------------------ app wiring
kvb_arg = app_src[app_src.index("KVB_ARG = {"):]
kvb_arg = kvb_arg[:kvb_arg.index("\n}\n")]
assert kvb_arg.count('"--keep-test-rows"') == 1, kvb_arg        # hitung_dw.py, KVB only

# ------------------------------------------------------------------ end to end
HDR_D = ["Source Name", "Sales UID", "Sales Name", "Sales Dept", "Currency", "Paid Date",
         "Transaction", "Payment Gateway", "USD", "Rate", "Ticket"]
HDR_W = ["Source Name", "Sales UID", "Sales Name", "Sales Dept", "Currency", "Payment Gateway",
         "Reference", "USD", "Rate", "Transaction", "Pay Channel", "Completed Date", "Status"]
REAL = (1, "Real Sales", "Real Dept")
TEST = (16, "Test data", "Test Department")


def buat(tmp):
    wb = Workbook()
    d = wb.active
    d.title = "D"
    d.append(HDR_D)
    d.append(["deposit", *REAL, "USDT", DT(2026, 10, 9), 23835.96, "USDT-BEP20", 23835.96, 1, "D1"])
    d.append(["deposit", *TEST, "USDT", DT(2026, 10, 9), 1000000, "USDT-BEP20", 1000000, 1, "D2"])
    d.append(["deposit", 7, "Real Name", "Test Department", "USD", DT(2026, 10, 9), 100,
              "Pay247-KH USD", 100, 1, "D3"])                                # dept only
    d.append(["deposit", *REAL, "USD", DT(2026, 10, 9), 40, "Pay247-KH USD", 40, 1, "D4"])
    w = wb.create_sheet("W")
    w.append(HDR_W)
    rows = [
        ("Withdrawal", *REAL, "USDT", "USDT-BEP20", 1001, 59093.44, 1, 59093.44, "USDT-BEP20", "finish"),
        ("Withdrawal", *REAL, "USDT", "Other", 1002, 50, 1, 50, "USDT-BEP20", "finish"),
        ("Withdrawal", *REAL, "USDT", "Other", 1003, 55, 1, 55, "USDT-ERC20", "finish"),
        ("Withdrawal", *REAL, "USDT", "Other", 1004, 7, 1, 7, "USDT-TRC20", "refuse"),
        ("Withdrawal", *TEST, "USDT", "Other", 1005, 10000, 1, 10000, "USDT-BEP20", "finish"),
        ("Rebate Withdrawal", *TEST, "USDT", "USDT-BEP20", 1006, 5, 1, 5, "USDT-BEP20", "success"),
        ("Wallet Withdrawal", *REAL, "USDT", "USDT-TRC20", 1007, 12, 1, 12, "USDT-TRC20", "success"),
    ]
    for (src, uid, nama, dept, cur, gw, ref, usd, rate, tx, pay, st) in rows:
        w.append([src, uid, nama, dept, cur, gw, ref, usd, rate, tx, pay, DT(2026, 10, 9), st])
    x = wb.create_sheet("Xero")
    x.append(["Date", "USDT", "USD"])
    x.append([DT(2026, 10, 9), 1, 1])
    fee = wb.create_sheet("D&W FEE")
    fee.append([None, None])
    fee.append([None, None])
    fee.append(["Currency", "Payment Gateway", "Deposit", "Withdrawal"])
    fee.append(["USDT", "Other", 0.0, 0.0])
    fee.append(["USDT", "USDT-TRC20", 0.0, 0.0])
    fee.append(["USD", "PAY247-KH", 0.01, 0.0])
    src = tmp / "dpm.xlsx"
    wb.save(src)
    opening = tmp / "open.txt"
    opening.write_text("BEP20\tUSDT\t1000\nOTHERS\tUSDT\t500\n", encoding="utf-8")
    return src, opening


def jalan(tmp, src, opening, *extra):
    dst = tmp / ("out" + "".join(e.strip("-")[:4] for e in extra) + ".xlsx")
    proc = subprocess.run([PY, str(HERE / "hitung_dw.py"), str(src), "-o", str(dst),
                           "--period", "2026-10", "--channel-opening", str(opening), *extra],
                          cwd=HERE, capture_output=True, text=True)
    assert proc.returncode == 0, proc.stdout[-3000:] + proc.stderr[-3000:]
    return load_workbook(dst, data_only=True), proc.stdout


def pcb(book):
    ws = book["Payment Channel Balance"]
    kol = [c.value for c in ws[1]]
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None or r[0].date() != datetime.date(2026, 10, 9):
            continue                                   # only the day with movements
        out[(r[1], r[2])] = dict(zip(kol, r))
    return out


def teks(ws):
    return "\n".join(" ".join(str(v) for v in r if v is not None)
                     for r in ws.iter_rows(values_only=True))


with tempfile.TemporaryDirectory() as t:
    tmp = Path(t)
    src, opening = buat(tmp)

    # ---------------- DPM (default)
    book, log = jalan(tmp, src, opening)
    p = pcb(book)
    usdt = {g.upper(): v for (c, g), v in p.items() if c == "USDT"}
    assert set(usdt) == {"BEP20", "OTHER", "TRC20"}, sorted(usdt)
    bep = usdt["BEP20"]
    assert abs(bep["Deposit"] - 23835.96) < 0.005, bep              # test 1,000,000 gone
    assert abs(bep["Withdrawal"] - 59093.44) < 0.005, bep           # no 'Other' rows inside
    assert abs(bep["Balance"] - (1000 + 23835.96 - 59093.44)) < 0.005, bep
    oth = usdt["OTHER"]
    assert abs(oth["Withdrawal"] - 105) < 0.005, oth                # 50 + 55; refuse + test out
    assert not oth["Deposit"], oth
    assert abs(oth["Balance"] - (500 - 105)) < 0.005, oth           # 'OTHERS' opening balance
    assert abs(usdt["TRC20"]["Withdrawal"] - 12) < 0.005, usdt["TRC20"]
    # USD: the Test Department deposit is out, the real one stays
    usd = [v for (c, g), v in p.items() if c == "USD"]
    assert len(usd) == 1 and abs(usd[0]["Deposit"] - 40) < 0.005, usd

    # D&W Report: the 'Other' line is 105, the BEP20 lines exclude the test rows
    rep = [r for r in book["D&W Report"].iter_rows(values_only=True)]
    lines = [(r[1], r[2]) for r in rep if r[1] in ("USDT-BEP20", "OTHER", "Other")]
    assert ("USDT-BEP20", 23835.96) in lines, lines
    assert ("USDT-BEP20", 59093.44) in lines, lines
    assert any(g in ("OTHER", "Other") and abs(v - 105) < 0.005 for g, v in lines), lines
    assert not any(v and v >= 10000 and g in ("OTHER", "Other") for g, v in lines), lines

    # D&W Detail has no test row at all
    det = teks(book["D&W Detail"])
    assert "Test data" not in det and "Test Department" not in det, det[:2000]

    # D / W: the calculated columns of test rows are blank
    for nama in ("D", "W"):
        ws = book[nama]
        kol = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            rec = dict(zip(kol, r))
            if rec["Sales Name"] == "Test data" or rec["Sales Dept"] == "Test Department":
                for c in ("Handling Fee", "Xero Rate", "Xero USD", "Forex Gain/Loss"):
                    assert rec.get(c) is None, (nama, rec)
            elif nama == "D" or rec["Status"] != "refuse":
                assert rec.get("Xero USD") is not None, (nama, rec)

    # MTOATD: USDT deposit of the day is the real one only
    mt = list(book["MTOATD"].iter_rows(values_only=True))
    i_kop = next(i for i, r in enumerate(mt) if r and r[0] == "日期：" and r[1] == DT(2026, 10, 9))
    i_usdt = list(mt[i_kop]).index("USDT")
    td = next(r for r in mt[i_kop:] if r and r[1] == "TD入金 (USD)")   # Paid Date basis
    assert abs(td[i_usdt] - 23835.96) < 0.005, td

    # Missing Data 0c lists the 4 test rows inside the month
    miss = teks(book["Missing Data"])
    assert "0c. TEST DATA - ROWS EXCLUDED AUTOMATICALLY" in miss, miss
    assert "TOTAL: 4 test rows excluded in this month." in miss, miss
    assert "Ref 1005" in miss and "Ref 1006" in miss and "USDT / Other" in miss, miss
    assert "Real Name / Test Department" in miss, miss
    leg = teks(book["Legend"])
    assert "v1.46" in leg and "Test data" in leg and "own channel row 'Other'" in leg, leg
    assert "DATA TES: 4 baris" in log, log[-2000:]
    book.close()

    # ---------------- KVB path: --keep-test-rows keeps the old behaviour for test rows
    book, log = jalan(tmp, src, opening, "--keep-test-rows")
    p = pcb(book)
    usdt = {g.upper(): v for (c, g), v in p.items() if c == "USDT"}
    assert abs(usdt["BEP20"]["Deposit"] - 1023835.96) < 0.005, usdt["BEP20"]
    assert abs(usdt["BEP20"]["Withdrawal"] - (59093.44 + 5)) < 0.005, usdt["BEP20"]
    assert abs(usdt["OTHER"]["Withdrawal"] - 10105) < 0.005, usdt["OTHER"]   # still own row
    miss = teks(book["Missing Data"])
    assert "0c. TEST DATA" not in miss, miss
    assert "DATA TES" not in log
    book.close()

print("ok test_dpm_usdt_other_test_data_v146")
