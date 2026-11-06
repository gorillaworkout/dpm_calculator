#!/usr/bin/env python3
"""Regression: USC Profit (USD) must be raw ÷ 100 (Bayu Nov 6 2026).

Live bug on Deals - Monthly Summary: Currency=USC was correct but
Profit / Commission / Fee / Swap (USD) equalled the raw amounts (e.g. 8.38
-> 8.38 instead of 0.0838). Cent accounts (Plus-10-cents) must always use
a flat ÷100 into the *(USD) columns — never an FX rate of 1.

Synthetic data only.
"""
import csv
import datetime
import io
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

from deal_segregator import proses

MT4_HEADER = [
    "Deal", "Login", "Open Time", "Type", "Symbol", "Volume",
    "Close Time", "Agent", "Commission", "Taxes", "Swap", "Profit", "Comment",
]


def _books(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            wb = load_workbook(io.BytesIO(z.read(name)), data_only=True)
            out[name] = {
                s: [tuple(r) for r in wb[s].iter_rows(values_only=True)]
                for s in wb.sheetnames
            }
            wb.close()
    return out


def _fx(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Query result"
    ws.append(["date", "Currency", "rate"])
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    wb.close()


def test_mt5_layout_usc_with_fx_rate_1_divides_usd_by_100():
    """MT5-layout Deals History, Currency=USC, FX rate 1 → USD cols = raw/100.

    This is the live shape Bayu hit: Currency stays USC, Profit stays the
    source amount, but Profit (USD) must be ÷100 (not equal to Profit).
    """
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        deals = tmp / "deals.csv"
        with deals.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Deal", "Login", "Time", "Type", "Entry", "Symbol", "Volume",
                "Commission", "Fee", "Swap", "Profit", "Currency",
            ])
            # Bayu sample-shaped amounts
            w.writerow([
                "1", "50001", "2026.10.15 01:00:00", "sell", "out", "EURGBP",
                "0.1", "0", "0", "0", "8.38", "USC",
            ])
            w.writerow([
                "2", "50001", "2026.10.16 01:00:00", "buy", "out", "XAUUSD",
                "0.2", "0", "0", "0", "361.69", "USC",
            ])
            w.writerow([
                "3", "50001", "2026.10.17 01:00:00", "sell", "out", "EURUSD",
                "0.3", "0", "0", "0", "-295.47", "USC",
            ])
        fx = tmp / "fx.xlsx"
        _fx(fx, [
            (datetime.date(2026, 10, 15), "USC", 1),
            (datetime.date(2026, 10, 16), "USC", 1),
            (datetime.date(2026, 10, 17), "USC", 1),
            (datetime.date(2026, 10, 15), "USD", 1),
            (datetime.date(2026, 10, 16), "USD", 1),
            (datetime.date(2026, 10, 17), "USD", 1),
        ])
        out = tmp / "out.zip"
        proses([deals, fx], out, allow_mt4=True)
        books = _books(out)
        monthly = books["Deals - Monthly Summary.xlsx"]["Monthly Summary"]
        header = monthly[0]
        assert "Profit (USD)" in header, header
        i_profit = header.index("Profit")
        i_cur = header.index("Currency")
        i_usd = header.index("Profit (USD)")
        i_sym = header.index("Symbol")
        by_sym = {r[i_sym]: r for r in monthly[1:]}
        # Profit column keeps source USC amounts; *(USD) are ÷100.
        assert by_sym["EURGBP"][i_cur] == "USC"
        assert by_sym["EURGBP"][i_profit] == 8.38
        assert by_sym["EURGBP"][i_usd] == 0.08  # round(8.38/100, 2)
        assert by_sym["XAUUSD"][i_profit] == 361.69
        assert by_sym["XAUUSD"][i_usd] == 3.62
        assert by_sym["EURUSD"][i_profit] == -295.47
        assert by_sym["EURUSD"][i_usd] == -2.95
        # Commission/Fee/Swap (USD) follow the same rule (zeros stay zero).
        for name in ("Commission (USD)", "Fee (USD)", "Swap (USD)"):
            i = header.index(name)
            assert by_sym["EURGBP"][i] == 0


def test_mt4_plus_10_cents_group_still_divides():
    """MT4 Raw Report + GROUP Plus-10-cents: sheet Profit already ÷100, USD copies."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        mt4 = tmp / "mt4.csv"
        with mt4.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(MT4_HEADER + ["", ""])
            w.writerow([
                "1", "50001", "2026.10.15 00:00:00", "sell", "EURGBP", "0.1",
                "2026.10.15 01:00:00", "0", "0", "0", "0", "8.38", "", "", "USC",
            ])
        peta = tmp / "accounts.csv"
        peta.write_text(
            "account,GROUP,account_type\n50001,real\\Plus-10-cents,Plus\n",
            encoding="utf-8",
        )
        fx = tmp / "fx.xlsx"
        _fx(fx, [
            (datetime.date(2026, 10, 15), "USC", 1),
            (datetime.date(2026, 10, 15), "USD", 1),
        ])
        out = tmp / "out.zip"
        proses([mt4, fx], out, allow_mt4=True, mt4_accounts=peta)
        books = _books(out)
        daily = books["Deals - Daily.xlsx"]["Daily"]
        header = daily[0]
        row = daily[1]
        assert row[header.index("Currency")] == "USC"
        assert row[header.index("Profit")] == 0.08
        assert row[header.index("Profit (USD)")] == 0.08


def test_account_type_plus_10_cents_without_group_is_usc():
    """account_type 'Plus-10-cents' alone (no GROUP col) must count as USC."""
    from deal_segregator import baca_peta_akun
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / "types.csv"
        p.write_text("account,account_type\n50001,Plus-10-cents\n", encoding="utf-8")
        assert baca_peta_akun(p) == {"50001": "USC"}


if __name__ == "__main__":
    test_mt5_layout_usc_with_fx_rate_1_divides_usd_by_100()
    print("OK mt5-layout USC FX rate 1 → USD ÷100")
    test_mt4_plus_10_cents_group_still_divides()
    print("OK mt4 Plus-10-cents group still ÷100")
    test_account_type_plus_10_cents_without_group_is_usc()
    print("OK account_type Plus-10-cents is USC")
    print("ALL test_deal_segregator_usc_divide_v142 OK")
