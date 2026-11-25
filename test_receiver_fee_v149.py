#!/usr/bin/env python3
"""Fund Transfer 手续费 on BOTH sides (v1.49).

Decisions of Bayu, 25 Nov 2026:
  1. The receiving-side 手续费 (收款 half, column N) is added to the RECEIVING channel's
     'Fund transfer charges' on Payment Channel Balance.
  2. It is in the receiving currency (收入币种).
  3. The receiving 金额 is GROSS: the fee is posted on top of it, nothing is counted twice.
  4. Cold Wallet is not special. J Wallet: the sending-side fee (H) is deducted when
     J Wallet sends, the receiving-side fee (N) when J Wallet receives, as a FEE row in
     the J Wallet ledger.
  5. Fees are signed: a negative fee is kept as it is (money back).
  6. One switch for the whole feature (hitung_dw.TRANSFER_FEE_V149 / --legacy-transfer-fee,
     fed per company by app.TRANSFER_FEE_V149).
USD(CNY) block: the incoming fee is converted with the receiving side's own rate.
A sheet whose N is empty must give exactly the v1.48 result.

All data here is synthetic.

Run: python3 test_receiver_fee_v149.py
"""
import datetime
import itertools
import subprocess
import sys
import tempfile
from pathlib import Path

import testkit

testkit.hermetic_env()

from openpyxl import Workbook, load_workbook  # noqa: E402

import hitung_dw as h  # noqa: E402

ROOT = Path(__file__).resolve().parent
DT = datetime.datetime
D = datetime.date


def near(a, b, eps=1e-6):
    return a is not None and abs(a - b) < eps


def _jalan(*args, cwd=ROOT):
    r = subprocess.run([sys.executable, *map(str, args)], cwd=cwd,
                       capture_output=True, text=True)
    assert r.returncode == 0, (args[0], r.stdout[-3000:], r.stderr[-3000:])
    return r.stdout


def _pcb(book):
    rows = list(book["Payment Channel Balance"].iter_rows(values_only=True))
    header = list(rows[0])
    out = []
    for r in rows[1:]:
        if all(v is None for v in r):
            continue
        item = dict(zip(header, r))
        if isinstance(item.get("Date"), datetime.datetime):
            item["Date"] = item["Date"].date()
        out.append(item)
    return out


def _row(pcb, cur, gw, tgl):
    hit = [r for r in pcb if r["Currency"] == cur and h.norm(r["Payment Gateway"]) == gw
           and r["Date"] == tgl]
    assert hit, (cur, gw, tgl, [(r["Currency"], r["Payment Gateway"], r["Date"]) for r in pcb])
    return hit[0]


def _last_balance(pcb, cur, gw):
    rows = [r for r in pcb if r["Currency"] == cur and h.norm(r["Payment Gateway"]) == gw
            and isinstance(r["Date"], D)]
    return max(rows, key=lambda r: r["Date"])["Balance"]


def _jw(book):
    return [r for r in book["J Wallet"].iter_rows(min_row=2, values_only=True)
            if any(v is not None for v in r)]


def _teks(ws):
    return "\n".join(str(v) for row in ws.iter_rows(values_only=True) for v in row
                     if isinstance(v, str))


def _dump(book, names):
    """Cell dump of the named sheets (generated time stamps masked)."""
    out = {}
    for name in names:
        rows = []
        for row in book[name].iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, str) and v.startswith("GENERATED "):
                    v = "GENERATED <time>"
                if isinstance(v, float):
                    v = round(v, 6)
                rows.append((cell.coordinate, v))
        out[name] = rows
    return out


def _same(a, b, what):
    assert a.keys() == b.keys(), what
    for name in a:
        if a[name] != b[name]:
            for left, right in itertools.zip_longest(a[name], b[name]):
                if left != right:
                    raise AssertionError(f"{what}: {name}: {left} != {right}")


# ---------------------------------------------------------------- the scenario
# (send date, send dept, channel, cur, amount, H fee, recv date, recv dept, recipient,
#  recv cur, recv amount, N fee)
N = None
SCENARIO = [
    # 1. channel -> J Wallet, sending fee only (N empty): H goes to TRC20, nothing on J
    (DT(2026, 11, 3), "KVB", "USDT-TRC20", "USDT", 10000, 20.13,
     DT(2026, 11, 3), "KVB", "J Wallet", "USDT", 10000, N),
    # 2. J Wallet -> Cold Wallet, both fees: H on the J Wallet ledger, N on Cold Wallet
    (DT(2026, 11, 5), "KVB", "J Wallet", "USDT", 5000, 33.40,
     DT(2026, 11, 5), "KVB", "Cold Wallet", "USDT", 5000, 50),
    # 3. channel -> channel (THB), both fees, same currency
    (DT(2026, 11, 10), "KVB", "PA", "THB", 1000, 10,
     DT(2026, 11, 10), "KVB", "1-2Pay", "THB", 1000, 15),
    # 4. channel -> J Wallet with a receiving fee: H on EP (CNY), N on the J Wallet ledger
    (DT(2026, 11, 12), "KVB", "EP", "CNY", 280000, 3,
     DT(2026, 11, 12), "KVB", "J Wallet", "USDT", 40000, 2.5),
    # 5. NEGATIVE fees: receiving fee -4 on 1-2Pay, and J Wallet sending fee -1.5
    (DT(2026, 11, 14), "KVB", "PA", "THB", 500, 5,
     DT(2026, 11, 14), "KVB", "1-2Pay", "THB", 500, -4),
    (DT(2026, 11, 15), "KVB", "J Wallet", "USDT", 300, -1.5,
     DT(2026, 11, 15), "KVB", "USDT-TRC20", "USDT", 300, N),
]

OPENING = [("USDT", "USDT-TRC20", 20000), ("USDT", "Cold Wallet", 0), ("THB", "PA", 5000),
           ("THB", "1-2Pay", 0), ("CNY", "EP", 0), ("USDT", "J Wallet", 1000)]


def _fee_cells(scenario, n_blank):
    return [row[:-1] + (None,) if n_blank else row for row in scenario]


def _dpm_source(path, scenario):
    """DPM-shaped workbook: D, W, Xero, D&W FEE, Fund Transfer Table (the DPM layout:
    header on row 3, 汇出 half B-H, 收款 half I-O), Opening Balance."""
    wb = Workbook()
    d = wb.active
    d.title = "D"
    d.append(["Source Name", "Currency", "Paid Date", "Transaction", "Payment Gateway",
              "USD", "Rate", "Ticket"])
    d.append(["Deposit", "USDT", DT(2026, 11, 1), 100, "USDT-TRC20", 100, 1, "D1"])
    w = wb.create_sheet("W")
    w.append(["Source Name", "Currency", "Completed Date", "Transaction", "Payment Gateway",
              "USD", "Rate", "Ticket", "Status"])
    w.append(["Withdrawal", "USDT", DT(2026, 11, 2), 10, "USDT-TRC20", 10, 1, "W1", "finish"])
    x = wb.create_sheet("Xero")
    x.append(["Date", "USDT", "THB", "CNY", "USD"])
    for day in range(1, 31):
        x.append([DT(2026, 11, day), 1, 32, 7, 1])
    f = wb.create_sheet("D&W FEE")
    f.append(["Currency", "Payment Gateway", "Deposit", "Withdrawal"])
    f.append(["USDT", "USDT-TRC20", 0.01, 0.01])
    ft = wb.create_sheet("Fund Transfer Table")
    ft.cell(2, 1, "序号")
    ft.cell(2, 2, "汇出")
    ft.cell(2, 9, "收款")
    hdr = {2: "汇款日期", 3: "汇出部门", 4: "支付渠道", 5: "汇出币种", 6: "Xero (USD)",
           7: "金额", 8: "手续费", 9: "收款日期", 10: "收款部门", 11: "收款人",
           12: "收入币种", 13: "金额", 14: "手续费", 15: "Xero (USD)"}
    for c, t in hdr.items():
        ft.cell(3, c, t)
    for i, (d1, dep1, ch, cur, amt, hfee, d2, dep2, rec, rcur, ramt, nfee) in enumerate(
            scenario, start=4):
        for c, v in zip((1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14),
                        (i - 3, d1, dep1, ch, cur, amt, hfee, d2, dep2, rec, rcur, ramt, nfee)):
            if v is not None:
                ft.cell(i, c, v)
    ob = wb.create_sheet("Opening Balance")
    ob.append(["As Of Date", "Currency", "Channel", "Opening Balance"])
    for cur, ch, val in OPENING:
        ob.append([DT(2026, 10, 31), cur, ch, val])
    wb.save(path)


# ============================================================ 1. parser reads both fees
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    _dpm_source(tmp / "dpm.xlsx", SCENARIO)
    wb = load_workbook(tmp / "dpm.xlsx")
    _nama, ftt = h.baca_fund_transfer(wb)
    assert len(ftt) == len(SCENARIO), len(ftt)
    by = {(t["tgl"], t["channel"]): t for t in ftt}
    t2 = by[(D(2026, 11, 5), "J Wallet")]
    assert t2["fee"] == 33.40 and t2["tujuan_fee"] == 50, t2        # H and N never mix
    t1 = by[(D(2026, 11, 3), "USDT-TRC20")]
    assert t1["fee"] == 20.13 and t1["tujuan_fee"] is None, t1
    assert by[(D(2026, 11, 14), "PA")]["tujuan_fee"] == -4          # sign kept
    assert by[(D(2026, 11, 15), "J Wallet")]["fee"] == -1.5
    wb.close()

# ============================================================ 2. DPM-layout run, switch ON
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "dpm.xlsx"
    _dpm_source(src, SCENARIO)
    out_on, out_off = tmp / "on.xlsx", tmp / "off.xlsx"
    log = _jalan("hitung_dw.py", src, "-o", out_on, "--period", "2026-11",
                 "--tanpa-extra-fees")
    assert "Transfer fee (v1.49)" in log and "3 baris 手续费" in log and "3 baris FEE" in log, log
    _jalan("hitung_dw.py", src, "-o", out_off, "--period", "2026-11",
           "--tanpa-extra-fees", "--legacy-transfer-fee")
    on = load_workbook(out_on, data_only=True)
    off = load_workbook(out_off, data_only=True)
    pcb, pcb0 = _pcb(on), _pcb(off)

    # ---- (1)+(2)+(3): receiving fee is a charge of the receiving channel, in its currency
    cold = _row(pcb, "USDT", "COLD WALLET", D(2026, 11, 5))
    assert near(cold["Fund transfer"], -5000) and near(cold["Fund transfer charges"], 50), cold
    assert near(cold["Balance"], 0 + 5000 - 50), cold              # gross 5000 in, 50 fee on top
    cold0 = _row(pcb0, "USDT", "COLD WALLET", D(2026, 11, 5))
    assert cold0["Fund transfer charges"] in (None, 0) and near(cold0["Balance"], 5000), cold0

    one = _row(pcb, "THB", "1-2PAY", D(2026, 11, 10))
    assert near(one["Fund transfer"], -1000) and near(one["Fund transfer charges"], 15), one
    pa = _row(pcb, "THB", "PA", D(2026, 11, 10))
    assert near(pa["Fund transfer"], 1000) and near(pa["Fund transfer charges"], 10), pa
    # the sending side is exactly as before (v1.48)
    assert near(_row(pcb0, "THB", "PA", D(2026, 11, 10))["Fund transfer charges"], 10)
    assert _row(pcb0, "THB", "1-2PAY", D(2026, 11, 10))["Fund transfer charges"] in (None, 0)

    # ---- (5) negative receiving fee kept as it is: it GIVES MONEY BACK
    neg = _row(pcb, "THB", "1-2PAY", D(2026, 11, 14))
    assert near(neg["Fund transfer"], -500) and near(neg["Fund transfer charges"], -4), neg
    assert near(neg["Balance"], one["Balance"] + 500 + 4), (neg["Balance"], one["Balance"])

    # ---- H is the SENDING channel's; J Wallet never shows up on this sheet
    ep = _row(pcb, "CNY", "EP", D(2026, 11, 12))
    assert near(ep["Fund transfer"], 280000) and near(ep["Fund transfer charges"], 3), ep
    assert not [r for r in pcb if "J WALLET" in h.norm(r["Payment Gateway"])]
    trc = _row(pcb, "USDT", "TRC20", D(2026, 11, 3))
    assert near(trc["Fund transfer"], 10000) and near(trc["Fund transfer charges"], 20.13), trc

    # final balances: opening - ft - ft charges (+ deposit/withdrawal where they exist)
    assert near(_last_balance(pcb, "USDT", "COLD WALLET"), 4950)
    assert near(_last_balance(pcb, "THB", "PA"), 5000 - 1010 - 505)
    assert near(_last_balance(pcb, "THB", "1-2PAY"), 1000 - 15 + 500 + 4)
    assert near(_last_balance(pcb, "CNY", "EP"), -280000 - 3)
    # the two sheets differ ONLY where a receiving-side fee exists
    assert near(_last_balance(pcb0, "USDT", "COLD WALLET"), 5000)
    assert near(_last_balance(pcb0, "THB", "1-2PAY"), 1000 + 500)

    # ---- J Wallet ledger: FEE rows under their transfer, balance drops with them
    led = _jw(on)
    kinds = [(r[0].date(), r[3]) for r in led if r[3] is not None]
    assert kinds == [
        (D(2026, 11, 3), "IN"),                       # N empty -> no fee row
        (D(2026, 11, 5), "OUT"), (D(2026, 11, 5), "FEE"),    # sending fee H = 33.40
        (D(2026, 11, 12), "IN"), (D(2026, 11, 12), "FEE"),   # receiving fee N = 2.5
        (D(2026, 11, 15), "OUT"), (D(2026, 11, 15), "FEE"),  # negative H = -1.5
    ], kinds
    bal = [r[7] for r in led]
    assert near(led[0][7], 1000) and led[0][1] == "Opening balance"
    # rows: open, IN 10000, OUT 5000, FEE 33.40, IN 40000, FEE 2.5, OUT 300, FEE -1.5
    assert [round(b, 2) for b in bal] == [1000, 11000, 6000, 5966.6, 45966.6, 45964.1,
                                          45664.1, 45665.6], bal
    fee1 = led[3]
    assert fee1[3] == "FEE" and near(fee1[6], -33.40) and fee1[5] is None, fee1   # out of wallet
    assert "sending side" in fee1[8], fee1
    fee2 = led[5]
    assert fee2[3] == "FEE" and near(fee2[6], -2.5) and "receiving side" in fee2[8], fee2
    fee3 = led[7]
    assert fee3[3] == "FEE" and near(fee3[5], 1.5) and fee3[6] is None, fee3      # money back
    assert "J Wallet" in _teks(on["Legend"]) and "v1.49" in _teks(on["Legend"])
    leg = _teks(on["Legend"])
    # receiving fees posted on channels: Cold Wallet 50, 1-2Pay 15, 1-2Pay -4 (the 2.5 on
    # J Wallet is a ledger FEE row, not a channel); J Wallet FEE rows: 33.40, 2.5, -1.5
    assert "3 receiving-side fee row(s) posted on channels, 3 J Wallet FEE row(s)" in leg, leg
    on.close()
    off.close()

    # ---- the switch OFF is v1.48: no receiving fee on channels, J Wallet ledger without fees
    off = load_workbook(out_off, data_only=True)
    assert [r[3] for r in _jw(off) if r[3] is not None] == ["IN", "OUT", "IN", "OUT"]
    assert "v1.49" not in _teks(off["Legend"])
    off.close()

# ============================================================ 3. EMPTY N = unchanged
# No receiving fee anywhere, and no sending fee on a J Wallet row: every sheet of the result
# must be identical with the switch on and off (the Legend gets no v1.49 note either).
SCEN_N_EMPTY = [row[:5] + ((None if row[2] == "J Wallet" else row[5]),) + row[6:11] + (N,)
                for row in SCENARIO]
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "dpm.xlsx"
    _dpm_source(src, SCEN_N_EMPTY)
    _jalan("hitung_dw.py", src, "-o", tmp / "on.xlsx", "--period", "2026-11",
           "--tanpa-extra-fees")
    _jalan("hitung_dw.py", src, "-o", tmp / "off.xlsx", "--period", "2026-11",
           "--tanpa-extra-fees", "--legacy-transfer-fee")
    on = load_workbook(tmp / "on.xlsx", data_only=True)
    off = load_workbook(tmp / "off.xlsx", data_only=True)
    assert on.sheetnames == off.sheetnames
    _same(_dump(on, on.sheetnames), _dump(off, off.sheetnames), "N empty, switch on vs off")
    assert "v1.49" not in _teks(on["Legend"])
    assert not [r for r in _jw(on) if r[3] == "FEE"]
    on.close()
    off.close()

    # N empty but a sending fee on a J Wallet row: Payment Channel Balance and every other
    # sheet are still identical; ONLY the J Wallet ledger (FEE rows) and the Legend note differ.
    src2 = tmp / "dpm2.xlsx"
    _dpm_source(src2, [row[:-1] + (N,) for row in SCENARIO])
    _jalan("hitung_dw.py", src2, "-o", tmp / "on2.xlsx", "--period", "2026-11",
           "--tanpa-extra-fees")
    _jalan("hitung_dw.py", src2, "-o", tmp / "off2.xlsx", "--period", "2026-11",
           "--tanpa-extra-fees", "--legacy-transfer-fee")
    on = load_workbook(tmp / "on2.xlsx", data_only=True)
    off = load_workbook(tmp / "off2.xlsx", data_only=True)
    same = [n for n in on.sheetnames if n not in ("J Wallet", "Legend")]
    _same(_dump(on, same), _dump(off, same), "N empty, J Wallet sending fee only")
    assert [r[3] for r in _jw(on) if r[3] == "FEE"] == ["FEE", "FEE"]     # H 33.40 and -1.5
    assert not [r for r in _jw(off) if r[3] == "FEE"]
    on.close()
    off.close()

    # switch OFF + N filled == switch OFF + N empty (N ignored, exactly as in v1.48)
    src3, src4 = tmp / "dpm3.xlsx", tmp / "dpm4.xlsx"
    _dpm_source(src3, SCENARIO)
    _dpm_source(src4, [row[:-1] + (N,) for row in SCENARIO])
    for src_, nm in ((src3, "a"), (src4, "b")):
        _jalan("hitung_dw.py", src_, "-o", tmp / f"leg_{nm}.xlsx", "--period", "2026-11",
               "--tanpa-extra-fees", "--legacy-transfer-fee")
    a = load_workbook(tmp / "leg_a.xlsx", data_only=True)
    b = load_workbook(tmp / "leg_b.xlsx", data_only=True)
    same = [n for n in a.sheetnames if n != "Fund Transfer Table"]    # that sheet holds N itself
    _same(_dump(a, same), _dump(b, same), "legacy switch ignores N")
    a.close()
    b.close()

# ============================================================ 4. KVB: USD(CNY) block
import app as app_mod  # noqa: E402


def _kvb_source(path, n_fee):
    wb = load_workbook(ROOT / "Template KVB Plus.xlsx")
    d = wb["D"]
    hd = [c.value for c in d[1]]
    vals = {"Source.Name": "11-2026 Deposit", "Paid Date": DT(2026, 11, 12),
            "Settlement Date": DT(2026, 11, 12), "method": "CNY", "ticket": "K1",
            "USD (CRM)": 100, "rate": 7, "transaction": 700,
            "receive_company_account": "ChipPay"}
    for k, v in vals.items():
        d.cell(2, hd.index(k) + 1, v)
    x = wb["Xero"]
    hx = [c.value for c in x[1]]
    for day in range(1, 31):
        x.cell(day + 1, 1, DT(2026, 11, day))
        for cur, v in (("CNY", 7), ("USD", 1), ("USDT", 1)):
            x.cell(day + 1, hx.index(cur) + 1, v)
    fee = wb["Handling Fee"]
    for r, (ch, cur) in enumerate((("ChipPay", "CNY"), ("EP", "CNY")), start=2):
        for i, v in enumerate((ch, cur, 0.01, 0.01), start=1):
            fee.cell(r, i, v)
    t = wb["Transfer"]
    ht = [c.value for c in t[2]]
    # H = sending 手续费 (first), N = receiving 手续费 (second 手续费 header)
    i_h = ht.index("手续费")
    i_n = ht.index("手续费", i_h + 1)
    row = {0: DT(2026, 11, 12), 1: "KVB", 2: "EP", 3: "CNY", 5: 7000, i_h: 10,
           8: DT(2026, 11, 12), 9: "KVB", 10: "ChipPay", 11: "CNY", 12: 7000}
    if n_fee is not None:
        row[i_n] = n_fee
    for c, v in row.items():
        t.cell(3, c + 1, v)
    wb.save(path)


def _kvb_run(tmp, n_fee, *extra):
    src = tmp / f"kvb_{n_fee}_{len(extra)}.xlsx"
    _kvb_source(src, n_fee)
    t, tpl0, tpl, out = (tmp / f"t{len(extra)}{n_fee}.xlsx", tmp / f"tpl0{len(extra)}{n_fee}.xlsx",
                         tmp / f"tpl{len(extra)}{n_fee}.xlsx", tmp / f"out{len(extra)}{n_fee}.xlsx")
    _jalan("siapkan_kvb.py", src, "-o", t)
    _jalan("buat_template.py", t, "-o", tpl0)
    _jalan("isi_template.py", t, "-t", tpl0, "--bulan", "2026-11",
           *app_mod.KVB_ARG["isi_template.py"], "-o", tpl)
    _jalan("hitung_dw.py", tpl, "--period", "2026-11", "-o", out,
           *app_mod.KVB_ARG["hitung_dw.py"], *extra)
    return load_workbook(out, data_only=True)


with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    book = _kvb_run(tmp, 5)
    pcb = _pcb(book)
    cp = _row(pcb, "CNY", "CHIPPAY", D(2026, 11, 12))
    ep = _row(pcb, "CNY", "EP", D(2026, 11, 12))
    assert near(cp["Fund transfer"], -7000) and near(cp["Fund transfer charges"], 5), cp
    assert near(ep["Fund transfer"], 7000) and near(ep["Fund transfer charges"], 10), ep
    # USD block: Xero USD = 7000 / 7 = 1000, rate 1/7 for BOTH sides; the incoming fee uses
    # the RECEIVING side's rate: 5 CNY / 7 = 0.714286 USD; the sending fee 10 / 7 = 1.428571
    usd_cp = _row(pcb, "USD", "CHIPPAY", D(2026, 11, 12))
    usd_ep = _row(pcb, "USD", "EP", D(2026, 11, 12))
    assert near(usd_cp["Fund transfer"], -1000, 1e-4), usd_cp
    assert near(usd_cp["Fund transfer charges"], 5 / 7, 1e-4), usd_cp
    assert near(usd_ep["Fund transfer"], 1000, 1e-4), usd_ep
    assert near(usd_ep["Fund transfer charges"], 10 / 7, 1e-4), usd_ep
    leg = _teks(book["Legend"])
    assert "receiving-side transfer fee (v1.49)" in leg
    assert "EP CNY out amount 7,000.00 fee 10.00" in leg, leg
    assert "ChipPay CNY in amount 7,000.00 fee 5.00 -> USD 1,000.00 fee 0.71" in leg, leg
    book.close()

    # the KVB switch off (what app.py passes when TRANSFER_FEE_V149['kvb'] is False)
    book = _kvb_run(tmp, 5, "--legacy-transfer-fee")
    pcb = _pcb(book)
    assert _row(pcb, "CNY", "CHIPPAY", D(2026, 11, 12))["Fund transfer charges"] in (None, 0)
    assert _row(pcb, "USD", "CHIPPAY", D(2026, 11, 12))["Fund transfer charges"] in (None, 0)
    assert near(_row(pcb, "USD", "EP", D(2026, 11, 12))["Fund transfer charges"], 10 / 7, 1e-4)
    assert "receiving-side transfer fee (v1.49)" not in _teks(book["Legend"])
    book.close()

    # N empty: the KVB result with the switch on equals the result with it off, every sheet
    on = _kvb_run(tmp, None)
    off = _kvb_run(tmp, None, "--legacy-transfer-fee")
    _same(_dump(on, on.sheetnames), _dump(off, off.sheetnames), "KVB, N empty")
    on.close()
    off.close()

# ============================================================ 5. one switch per company
assert app_mod.TRANSFER_FEE_V149 == {"dpm": True, "kvb": True}, app_mod.TRANSFER_FEE_V149
app_src = (ROOT / "app.py").read_text(encoding="utf-8")
assert app_src.count('"--legacy-transfer-fee"') == 1, "exactly one place passes the flag"
assert 'if script == "hitung_dw.py" and not TRANSFER_FEE_V149.get(company, True):' in app_src
kvb_arg = app_src[app_src.index("KVB_ARG = {"):]
kvb_arg = kvb_arg[:kvb_arg.index("\n}\n")]
assert "--legacy-transfer-fee" not in kvb_arg                 # not hard-wired for KVB
assert h.TRANSFER_FEE_V149 is True                            # default: both sides count

# ============================================================ 6. version
ver = (ROOT / "VERSION").read_text(encoding="utf-8")
assert "version=1.49" in ver and "build=2026-11-25" in ver, ver
assert "v1.49" in (ROOT / "README.md").read_text(encoding="utf-8")

print("test_receiver_fee_v149: ok")
