import csv
import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

import deal_segregator
from deal_segregator import proses


HEADERS = [
    "Deal", "Login", "Group", "Country", "Time", "Type", "Entry", "Symbol",
    "Volume", "Commission", "Fee", "Swap", "Profit", "Currency",
]
RESULT_HEADERS = [
    "Login", "Country", "Desk", "Date", "Type", "Symbol", "Deals", "Volume",
    "Commission", "Fee", "Swap", "Profit", "Currency",
]


def row(deal, login="100", group=r"real\DPMKT-15", country="MY",
        time="2026.07.01 01:02:03.004", type_="buy", entry="out",
        symbol="EURUSD", volume="1.25", commission="-1", fee="-0.1",
        swap="-0.2", profit="10", currency="USD"):
    return [deal, login, group, country, time, type_, entry, symbol, volume,
            commission, fee, swap, profit, currency]


def write_csv(path, rows):
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(HEADERS)
        writer.writerows(rows)


def values(ws):
    return list(ws.iter_rows(values_only=True))


def expect_error(rows, *parts):
    with TemporaryDirectory() as tmp:
        src = Path(tmp) / "bad.csv"
        write_csv(src, rows)
        try:
            proses([src], Path(tmp) / "out.xlsx")
        except SystemExit as exc:
            message = str(exc)
            for part in parts:
                assert part in message, (part, message)
        else:
            raise AssertionError("invalid input was accepted")


# Portable synthetic fixture exercises all financial and grouping behavior.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    first = tmp / "july.csv"
    second = tmp / "august.csv"
    write_csv(first, [
        row("1"),
        row("2", volume="0.75", commission="-2", fee="-0.2", swap="-0.3", profit="5"),
        row("3", time="2026.07.02 04:05:06", volume="1", commission="-1", fee="-0.3", swap="-0.4", profit="20"),
        row("4", login="200", group=r"real\DPMKT-20", country="SG",
            time="2026.08.03 00:00:00", type_="sell", symbol="XAUUSD",
            volume="2", commission="-4", fee="-0.4", swap="-0.5", profit="30"),
        row("ignored", entry="in", profit="999"),
    ])
    write_csv(second, [
        row("2", time="2026.08.04 00:00:00", profit="9999"),  # duplicate ignored
        row("5", time="2026.08.04 00:00:00", volume="0.5", commission="-0.5",
            fee="-0.05", swap="-0.1", profit="7"),
    ])
    out = tmp / "hasil.xlsx"
    summary = proses([first, second], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        assert wb.sheetnames == ["Daily", "Monthly Summary", "Verifikasi"]
        daily = values(wb["Daily"])
        monthly = values(wb["Monthly Summary"])
        assert daily[0] == tuple(RESULT_HEADERS)
        assert monthly[0] == tuple("Month" if h == "Date" else h for h in RESULT_HEADERS)
        assert daily[1:] == [
            ("100", "MY", "DPMKT-15", datetime.datetime(2026, 7, 1), "buy", "EURUSD", 2, 2, -3, -0.3, -0.5, 15, "USD"),
            ("100", "MY", "DPMKT-15", datetime.datetime(2026, 7, 2), "buy", "EURUSD", 1, 1, -1, -0.3, -0.4, 20, "USD"),
            ("100", "MY", "DPMKT-15", datetime.datetime(2026, 8, 4), "buy", "EURUSD", 1, 0.5, -0.5, -0.05, -0.1, 7, "USD"),
            ("200", "SG", "DPMKT-20", datetime.datetime(2026, 8, 3), "sell", "XAUUSD", 1, 2, -4, -0.4, -0.5, 30, "USD"),
        ]
        assert monthly[1:] == [
            ("100", "MY", "DPMKT-15", datetime.datetime(2026, 7, 1), "buy", "EURUSD", 3, 3, -4, -0.6, -0.9, 35, "USD"),
            ("100", "MY", "DPMKT-15", datetime.datetime(2026, 8, 1), "buy", "EURUSD", 1, 0.5, -0.5, -0.05, -0.1, 7, "USD"),
            ("200", "SG", "DPMKT-20", datetime.datetime(2026, 8, 1), "sell", "XAUUSD", 1, 2, -4, -0.4, -0.5, 30, "USD"),
        ]
        verification = dict(values(wb["Verifikasi"])[2:])
        assert verification["Total baris (semua file)"] == 7
        assert verification["Entry = out (dipakai)"] == 6
        assert verification["Entry = in (dibuang)"] == 1
        assert verification["Duplikat antar-file (Deal ID sama, dibuang)"] == 1
        assert verification["Baris unik dipakai"] == 5
        assert verification["Baris pada Daily (Login+Date+Type+Symbol)"] == 4
        assert verification["Baris pada Monthly Summary (Login+Month+Type+Symbol)"] == 3
        assert verification["Login unik"] == 2
        assert verification["Country terisi"] == 5
        assert verification["Country kosong"] == 0
        assert verification["Desk unik"] == 2
        assert verification["Jumlah Profit (semua baris out)"] == 72
        assert summary == {"file_masuk": 2, "total": 7, "dipakai": 5,
                           "harian": 4, "bulanan": 3, "login_unik": 2}
    finally:
        wb.close()

# Trust-boundary diagnostics include source location and offending column.
expect_error([row("1", volume="not-a-number")], "bad.csv", "baris 2", "Volume")
expect_error([row("1", fee="")], "bad.csv", "baris 2", "Fee")
expect_error([row("1", profit="NaN")], "bad.csv", "baris 2", "Profit", "NaN")
expect_error([row("1", swap="inf")], "bad.csv", "baris 2", "Swap", "inf")
expect_error([row("1", time="")], "bad.csv", "baris 2", "Time")
expect_error([row("1", time="2026.02.30 00:00:00")], "bad.csv", "baris 2", "Time")
expect_error([row("1", time="2026.07.01 25:00:00")], "bad.csv", "baris 2", "Time")
expect_error([row("1", time="2026.07.01 garbage")], "bad.csv", "baris 2", "Time")
expect_error([row("1", time="yesterday")], "bad.csv", "baris 2", "Time")
expect_error([["1", "100"]], "bad.csv", "baris 2", "14 kolom", "2 kolom")

# A rejected XLSX must not retain an open read-only workbook.
with TemporaryDirectory() as tmp:
    src = Path(tmp) / "missing-columns.xlsx"
    workbook = Workbook()
    workbook.active.append(["Deal"])
    workbook.save(src)
    workbook.close()
    real_load_workbook = deal_segregator.load_workbook
    opened = []
    def tracking_load_workbook(*args, **kwargs):
        opened_workbook = real_load_workbook(*args, **kwargs)
        opened.append(opened_workbook)
        return opened_workbook
    deal_segregator.load_workbook = tracking_load_workbook
    try:
        try:
            proses([src], Path(tmp) / "out.xlsx")
        except SystemExit:
            pass
        else:
            raise AssertionError("missing columns were accepted")
        assert opened and opened[0]._archive.fp is None
    finally:
        deal_segregator.load_workbook = real_load_workbook

print("OK deal segregator engine")
