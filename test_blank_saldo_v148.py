#!/usr/bin/env python3
"""v1.48: a blank opening-balance cell is 0, and trailing empty Excel columns
do not shift the row.

When a tab-split Payment Channel Balance paste had more than 3 parts, the
parser dropped every empty cell. A row with a blank Balance plus extra trailing
columns (how Excel often pastes a selected range), e.g. 'RM Pay\\tJPY\\t\\t',
collapsed to two parts and was read as currency 'RM Pay', saldo 'JPY', with the
channel carried down from the previous row (SojuPay). That raised:

    Saldo channel SojuPay RM Pay bukan angka yang bisa dibaca: 'JPY'

Synthetic data only (same channel/currency names as a real paste, made-up
amounts, no client figures).    python3 test_blank_saldo_v148.py
"""
import sys
from pathlib import Path

import hitung_dw as h

HERE = Path(__file__).resolve().parent
assert Path(h.__file__).resolve().parent == HERE

# Channel / currency pairs in the same order as the team's paste. Amounts are
# made-up (None = the Balance cell is left blank).
ROWS = [
    ("OCPAY", "VND", 110.25),
    ("BeckPay", "VND", 220.50),
    ("Rivo", "VND", 330.75),
    ("Rivo", "THB", 440.00),
    ("PA", "ZAR", 550.25),
    ("PA", "BRL", 660.50),
    ("PA", "MXN", 770.75),
    ("PA", "KHR-USD", 880.00),
    ("PA", "PHP", 990.25),
    ("PA", "THB", 1010.50),
    ("BuziPay", "USD", 1111.75),
    ("DUPOIN GLOBAL INC", "USD", 1222.00),
    ("Flowin-JCB", "USD", 1333.25),
    ("GDPay", "USD", 1444.50),
    ("GDPay", "EUR", 1555.75),
    ("MONETIX", "Binance USD", 1666.00),
    ("MONETIX", "EGP", 1777.25),
    ("MONETIX", "NGN", 1888.50),
    ("MONETIX", "KES", 1999.75),
    ("MONETIX", "ZAR", 2110.00),
    ("MONETIX", "KHR-USD", -12.50),
    ("BeckPay", "INR", 13.25),
    ("AXIS BANK", "INR", 2222.00),
    ("KOTAK BANK", "INR", 2333.25),
    ("KOTAK BANK2", "INR", -144.00),
    ("INDUSIND BANK", "INR", 2555.75),
    ("1-2Pay", "THB", 2666.00),
    ("ZENPAY", "THB", 2777.25),
    ("Payport", "UZS", 2888.50),
    ("Vexora", "EGP", 2999.75),
    ("Vexora", "NGN", 3110.00),
    ("Pay247", "KHR USD", 3222.25),
    ("Pay247", "INR", 3333.50),
    ("Pay247", "SGD", 3444.75),
    ("Pay247", "PKR", 3555.00),
    ("Pay247", "LKR", 3666.25),
    ("Thunder X Pay", "LAK", 3777.50),
    ("Thunder X Pay", "LAK USD", 3888.75),
    ("Peska", "JPY", 3999.00),
    ("Skymarkets-JPY", "JPY", 4110.25),
    ("External GBP Bank", "GBP", 4222.50),
    ("USDT-TRC20", "USDT", 4333.75),
    ("USDT-ERC20", "USDT", 4444.00),
    ("USDT-BEP20", "USDT", 4555.25),
    ("USDC-ERC20", "USDT", 4666.50),
    ("USDC-BEP20", "USDT", 4777.75),
    ("77Pay", "VND", 0.0),
    ("Novolink", "VND", -1888.50),
    ("PA", "VND", 4999.00),
    ("PA", "INR", 0.25),
    ("SHUNFA PAY", "PHP", 0.50),
    ("MoPay - USDT", "USDT", -0.75),
    ("Mopay - VND", "VND", -16.25),
    ("24Pay", "UZS", 5110.00),
    ("Trustlink", "INR", 0.0),
    ("MYP", "INR", 5222.25),
    ("MYP", "THB", 5333.50),
    ("SojuPay", "KRW", 12000.25),
    ("RM Pay", "JPY", None),
    ("OnlinePay", "USD", None),
    ("Cash Deposit", "AED", 5444.75),
    ("VPAY", "THB", -17.50),
    ("OVP", "THB", 5555.00),
    ("BitPayz", "THB", 5666.25),
    ("BerryPay", "THB", 5777.50),
    ("NEPay", "THB", None),
    ("Thunder X Pay", "KHR", None),
    ("Thunder X Pay", "KHR USD", 0.25),
    ("FuturePay", "ZAR", 0.50),
    ("MONETIX", "TRY", None),
    ("MONETIX", "PKR", 0.0),
    ("MONETIX", "KRW", -0.75),
    ("MONETIX", "LAK-USD", 5888.75),
    ("MONETIX", "INR", 5999.00),
    ("MONETIX", "UZS", -1999.25),
]

BLANK = {("RM Pay", "JPY"), ("OnlinePay", "USD"), ("NEPay", "THB"),
         ("Thunder X Pay", "KHR"), ("MONETIX", "TRY")}
assert {r[:2] for r in ROWS if r[2] is None} == BLANK


def format_amt(v):
    """Write a made-up amount the way a paste usually looks (minus, 2 decimals)."""
    if v is None:
        return ""
    return f"{v:.2f}"


def paste(rows, trailing="", omit_blank_third=False):
    """Build a tab-separated paste.

    trailing: extra empty columns appended to every row that has a Balance,
              and to blank-balance rows unless omit_blank_third is used.
    omit_blank_third: blank-balance rows are only 'Channel\\tCurrency'
                      (no third cell at all).
    """
    lines = []
    for ch, cur, v in rows:
        if v is None:
            if omit_blank_third:
                lines.append(f"{ch}\t{cur}")
            else:
                lines.append(f"{ch}\t{cur}\t{trailing}")
        else:
            lines.append(f"{ch}\t{cur}\t{format_amt(v)}{trailing}")
    return "\n".join(lines)


def lookup(out, ch, cur):
    key = (h.currency_shadow_key(cur) or h.norm(cur), h.kunci_channel(cur, ch))
    assert key in out, (ch, cur, key, sorted(out))
    return out[key]


def assert_rows(out, n, rows):
    assert n == len(rows), (n, len(rows))
    assert len(out) == len(rows), (len(out), len(rows), sorted(out))
    for ch, cur, v in rows:
        got = lookup(out, ch, cur)
        want = 0.0 if v is None else float(v)
        assert got == want, (ch, cur, got, want)


def parse(teks):
    return h.baca_tabel_saldo_channel(teks)


# ------------------------------------------------------------------ repro
teks = "SojuPay\tKRW\t12000.25\nRM Pay\tJPY\t\t"
out, n = parse(teks)
assert n == 2, n
assert lookup(out, "SojuPay", "KRW") == 12000.25
assert lookup(out, "RM Pay", "JPY") == 0.0
assert ("RM PAY", h.kunci_channel("RM PAY", "SojuPay")) not in out
print("OK repro trailing empty columns (the v1.47 error case)")

# same shape, no extra trailing columns (third cell present but empty)
out, n = parse("SojuPay\tKRW\t12000.25\nRM Pay\tJPY\t")
assert n == 2 and lookup(out, "RM Pay", "JPY") == 0.0
assert lookup(out, "SojuPay", "KRW") == 12000.25
print("OK blank saldo without extra trailing columns")

# two columns only: Channel<TAB>Currency, no third cell
out, n = parse("SojuPay\tKRW\t12000.25\nRM Pay\tJPY")
assert n == 2 and lookup(out, "RM Pay", "JPY") == 0.0
assert lookup(out, "SojuPay", "KRW") == 12000.25
print("OK blank saldo as two columns (no third cell)")

# ------------------------------------------------------------------ merged channel
# Leading empty channel cell = carry down, with and without trailing empties,
# and with a blank saldo on a carried-down row.
merged = (
    "PA\tZAR\t1100.00\n"
    "\tBRL\t2200.00\n"
    "\tMXN\t3300.00\t\t\n"
    "\tPHP\t\t\n"
)
out, n = parse(merged)
assert n == 4, (n, out)
assert lookup(out, "PA", "ZAR") == 1100.00
assert lookup(out, "PA", "BRL") == 2200.00
assert lookup(out, "PA", "MXN") == 3300.00
assert lookup(out, "PA", "PHP") == 0.0
print("OK merged-channel carry-down (incl. trailing empties and blank saldo)")

# the two shapes called out in the brief
out, n = parse("PA\tZAR\t1100.00\n\tBRL\t2200.00")
assert n == 2 and lookup(out, "PA", "BRL") == 2200.00
out, n = parse("PA\tZAR\t1100.00\n\tBRL\t2200.00\t\t")
assert n == 2 and lookup(out, "PA", "BRL") == 2200.00
print("OK '\\tBRL\\t5399.16' and '\\tBRL\\t5399.16\\t\\t' shapes")

# blank channel + currency + blank saldo, including the form whose .strip()
# would otherwise leave only the currency (no tab left)
for row in ("\tPHP\t\t", "\tPHP\t", "\tPHP"):
    out, n = parse("PA\tZAR\t1100.00\n" + row)
    assert n == 2, (row, n, out)
    assert lookup(out, "PA", "ZAR") == 1100.00
    assert lookup(out, "PA", "PHP") == 0.0, (row, out)
print("OK blank channel + currency + blank saldo")

# ------------------------------------------------------------------ old empty-dropping preserved
# Extra leading empty column on a 3-value row (selecting one column too far
# to the left). The old 'drop empties when len > 3' handled this; keep it.
out, n = parse("\tSojuPay\tKRW\t12000.25")
assert n == 1 and lookup(out, "SojuPay", "KRW") == 12000.25
# spacer in the middle of 4 parts
out, n = parse("SojuPay\t\tKRW\t12000.25")
assert n == 1 and lookup(out, "SojuPay", "KRW") == 12000.25
print("OK old empty-dropping for extra leading / middle columns")

# 2-col combined channel+currency still does not inherit the previous channel
out, n = parse("Beckpay\tVND\t100\nExamplePay LKR\t250\n\tUSD\t50")
assert n == 3
assert lookup(out, "Beckpay", "VND") == 100
assert lookup(out, "ExamplePay", "LKR") == 250
assert lookup(out, "ExamplePay", "USD") == 50
print("OK combined channel+currency 2-col row (regression 17 Sep 2026)")

# dash still means 0; a typo in a 3-col amount still errors
out, n = parse("PA\tBRL\t-")
assert n == 1 and lookup(out, "PA", "BRL") == 0.0
try:
    parse("PA\tVND\t12abc34\n")
    raise AssertionError("must exit")
except SystemExit as e:
    pesan = str(e.code)
    assert "'12abc34'" in pesan, pesan
    assert "Saldo channel PA VND" in pesan, pesan
print("OK dash=0 and unreadable amount still errors")

# ------------------------------------------------------------------ full paste, both shapes
for label, teks in (
    ("with trailing empty columns on blank rows", paste(ROWS, trailing="\t\t")),
    ("without trailing empty columns", paste(ROWS, trailing="")),
    ("blank rows as two columns only", paste(ROWS, trailing="", omit_blank_third=True)),
):
    out, n = parse(teks)
    assert_rows(out, n, ROWS)
    for ch, cur in BLANK:
        assert lookup(out, ch, cur) == 0.0, (label, ch, cur)
    assert lookup(out, "SojuPay", "KRW") == 12000.25
    assert lookup(out, "Cash Deposit", "AED") == 5444.75
    assert lookup(out, "Thunder X Pay", "KHR USD") == 0.25
    assert lookup(out, "MONETIX", "PKR") == 0.0
    print(f"OK full paste {label}: {n} rows, blank balances = 0, no error")

print("OK blank saldo v1.48")
