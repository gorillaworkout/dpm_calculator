import csv
import datetime
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

import deal_segregator


def test_zip_writer_does_not_buffer_workbooks_in_memory():
    original = zipfile.ZipFile.writestr
    def forbidden(archive, name, data, *args, **kwargs):
        if str(name).startswith("Deals - "):
            raise AssertionError("large workbooks must be added from disk, not bytes")
        return original(archive, name, data, *args, **kwargs)
    zipfile.ZipFile.writestr = forbidden
    row = {"Login": "1", "Periode": datetime.date(2026, 9, 1), "Type": "buy",
           "Symbol": "EURUSD", "Deals": 1, "Volume": 1.0,
           "Commission": -1.0, "Fee": 0.0, "Swap": 0.0,
           "Profit": 2.0, "Currency": "USD", "_kurs_kurang": 0}
    try:
        with TemporaryDirectory() as folder:
            out = Path(folder) / "result.zip"
            deal_segregator._tulis_zip([row], [row], out, {}, False, set())
            with zipfile.ZipFile(out) as archive:
                assert archive.namelist() == ["Deals - Daily.xlsx", "Deals - Monthly Summary.xlsx"]
    finally:
        zipfile.ZipFile.writestr = original
from deal_segregator import proses


HEADERS = [
    "Deal", "Position", "Login", "Group", "Country", "Time", "Type", "Entry",
    "Symbol", "Volume", "Commission", "Fee", "Swap", "Profit", "Currency",
]
RESULT_HEADERS = [
    "Login", "Date", "Type", "Symbol", "Deals", "Volume",
    "Commission", "Fee", "Swap", "Profit", "Currency",
]


def row(deal, position=None, login="100", group=r"real\DPMKT-15", country="MY",
        time="2026.07.01 01:02:03.004", type_="buy", entry="out",
        symbol="EURUSD", volume="1.25", commission="-1", fee="-0.1",
        swap="-0.2", profit="10", currency="USD"):
    # Kalau tidak diberikan, Position default UNIK per baris (= Deal-nya sendiri)
    # supaya baris di fixture yang sudah ada TIDAK saling ketemu lewat pencarian
    # Position (lihat _kumpulkan_komisi_in/_normalisasi_satu_file) -- angka yang
    # sudah diverifikasi di bawah tetap berlaku apa adanya. Tes gabungan in/out
    # yang SENGAJA berbagi Position ada di blok terpisah di bawah.
    if position is None:
        position = deal
    return [deal, position, login, group, country, time, type_, entry, symbol,
            volume, commission, fee, swap, profit, currency]


def write_csv(path, rows):
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(HEADERS)
        writer.writerows(rows)


def values(ws):
    return list(ws.iter_rows(values_only=True))


# Exact 22-column DPM MT5 export: Group/Country are intentionally absent.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "July_2026_-_Deals_History_DPM_MT5.csv", tmp / "out.xlsx"
    headers = ["Deal", "ID", "Order", "Position", "Login", "Name", "Time", "Type",
               "Entry", "Symbol", "Volume", "Price", "S / L", "T / P", "Reason",
               "Commission", "Fee", "Swap", "Profit", "Dealer", "Currency", "Comment"]
    data = ["1", "id", "o", "P1", "100", "Client", "2026.07.01 00:00:00",
            "buy", "out", "EURUSD", "1", "1.1", "", "", "client", "-2", "-1",
            "-3", "20", "", "USD", ""]
    with src.open("w", encoding="utf-8", newline="") as f:
        csv.writer(f).writerows([headers, data])
    proses([src], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        assert values(wb["Daily"])[0] == tuple(RESULT_HEADERS)
        assert values(wb["Daily"])[1][0] == "100"
        assert "Country" not in RESULT_HEADERS and "Desk" not in RESULT_HEADERS
    finally:
        wb.close()


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
            ("100", datetime.datetime(2026, 7, 1), "buy", "EURUSD", 2, 2, -3, -0.3, -0.5, 15, "USD"),
            ("100", datetime.datetime(2026, 7, 2), "buy", "EURUSD", 1, 1, -1, -0.3, -0.4, 20, "USD"),
            ("100", datetime.datetime(2026, 8, 4), "buy", "EURUSD", 1, 0.5, -0.5, -0.05, -0.1, 7, "USD"),
            ("200", datetime.datetime(2026, 8, 3), "sell", "XAUUSD", 1, 2, -4, -0.4, -0.5, 30, "USD"),
        ]
        assert monthly[1:] == [
            ("100", datetime.datetime(2026, 7, 1), "buy", "EURUSD", 3, 3, -4, -0.6, -0.9, 35, "USD"),
            ("100", datetime.datetime(2026, 8, 1), "buy", "EURUSD", 1, 0.5, -0.5, -0.05, -0.1, 7, "USD"),
            ("200", datetime.datetime(2026, 8, 1), "sell", "XAUUSD", 1, 2, -4, -0.4, -0.5, 30, "USD"),
        ]
        verification = dict(values(wb["Verifikasi"])[2:])
        assert wb["Verifikasi"]["A1"].value == "Summary - Deal Segregator"
        assert verification["Total source rows (all files)"] == 7
        assert verification["Entry = out (before deduplication)"] == 6
        assert verification["Entry = in (merged onto matching Out by Position)"] == \
            "1 seen, 0 rows merged onto a matching Out row"
        assert verification["In-row commission with no matching Out in this upload (Position still open?)"] == \
            "1 positions, 1 'in' rows"
        assert verification["Duplicate non-empty Deal IDs (dropped)"] == 1
        assert verification["Unique rows kept"] == 5
        # The label must state every grouping dimension actually used.
        assert verification["Daily output rows (Login+Date+Type+Symbol+Currency)"] == 4
        assert verification["Monthly Summary output rows (Login+Month+Type+Symbol+Currency)"] == 3
        assert verification["Unique logins"] == 2
        assert "Rows with Country" not in verification
        assert "Rows without Country" not in verification
        assert "Unique desks" not in verification
        assert verification["Total Profit (all kept rows)"] == 72
        assert summary == {"file_masuk": 2, "total": 7, "dipakai": 5,
                           "harian": 4, "bulanan": 3, "login_unik": 2}
    finally:
        wb.close()

# Entry='in' Commission/Fee/Swap merge onto the matching Entry='out' row by
# Position (dikonfirmasi user 24 Sep 2026, opsi "merge onto the closing row"):
# some brokers only record Commission on the opening leg, so dropping 'in'
# rows outright silently loses real charges. Volume/Deals must stay exactly
# as the 'out' row alone -- 'in' and 'out' of the same position almost always
# report the SAME Volume, so counting both would double it.
# Each Position below uses its OWN Symbol so the three scenarios land in
# three distinct, unambiguous output rows (grouping is by Symbol among other
# fields, so reusing a Symbol across positions would merge them together).
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "merge.csv", tmp / "merge.xlsx"
    write_csv(src, [
        # Simple case: one 'in' with real Commission, one 'out' with none of
        # its own -- merged Commission must equal the 'in' row's alone.
        row("1", position="P1", entry="in", commission="-2", fee="0", swap="0", profit="0"),
        row("2", position="P1", entry="out", symbol="XAUUSD", commission="0", fee="0", swap="0", profit="10"),
        # Partial close: ONE 'in' shared by TWO 'out' rows (different Symbols,
        # so they land in different output rows) -- the 'in' Commission must
        # land on only the FIRST 'out' encountered in the file, never both.
        row("3", position="P2", entry="in", commission="-5", fee="-0.5", swap="0", profit="0"),
        row("4", position="P2", entry="out", symbol="EURUSD", commission="0", fee="0", swap="0", profit="20"),
        row("5", position="P2", entry="out", symbol="GBPUSD", commission="0", fee="0", swap="0", profit="30"),
        # Position never closed in this upload -- its Commission must be
        # reported as unmatched, not silently dropped.
        row("6", position="P3", entry="in", commission="-7", fee="0", swap="0", profit="0"),
    ])
    summary = proses([src], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        by_symbol = {r[3]: r for r in values(wb["Daily"])[1:]}
        assert by_symbol["XAUUSD"][6] == -2, "P1: merged Commission must be exactly the 'in' row's -2"
        assert by_symbol["XAUUSD"][4] == 1, "Deals must stay 1, not 2"
        assert by_symbol["XAUUSD"][5] == 1.25, "Volume must not double"
        assert by_symbol["EURUSD"][6] == -5 and by_symbol["EURUSD"][7] == -0.5, \
            "P2 first Out (EURUSD, file order) must claim the whole 'in' Commission/Fee"
        assert by_symbol["GBPUSD"][6] == 0 and by_symbol["GBPUSD"][7] == 0, \
            "P2 second Out (GBPUSD, partial close) must NOT also get the 'in' Commission"
        verification = dict(values(wb["Verifikasi"])[2:])
        assert verification["Entry = in (merged onto matching Out by Position)"] == \
            "3 seen, 2 rows merged onto a matching Out row"
        assert verification["In-row commission with no matching Out in this upload (Position still open?)"] == \
            "1 positions, 1 'in' rows"
    finally:
        wb.close()

# A duplicate Out Deal must not become the opening-cost target and then be
# discarded. Target selection must use the same first-seen Deal dedup rule.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    opening, kept, duplicate = tmp / "opening.csv", tmp / "kept.csv", tmp / "duplicate.csv"
    write_csv(opening, [row("in", position="PD", entry="in", time="2026.01.01 00:00:00",
                            commission="-9", fee="0", swap="0", profit="0")])
    write_csv(kept, [row("same-deal", position="PD", entry="out", time="2026.02.01 00:00:00",
                         symbol="KEPT", commission="0", fee="0", swap="0", profit="2")])
    write_csv(duplicate, [row("same-deal", position="PD", entry="out", time="2026.01.15 00:00:00",
                              symbol="DUPLICATE", commission="0", fee="0", swap="0", profit="1")])
    out = tmp / "dedup-target.xlsx"
    proses([opening, kept, duplicate], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        rows = values(wb["Daily"])[1:]
        assert len(rows) == 1
        assert rows[0][3] == "KEPT" and rows[0][6] == -9, rows
    finally:
        wb.close()

# Upload order must not decide which partial close receives the opening costs.
# The earliest Out by Time owns them, even when its file is uploaded second.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    opening, later, earlier = tmp / "opening.csv", tmp / "later.csv", tmp / "earlier.csv"
    write_csv(opening, [
        row("10", position="PX", entry="in", time="2026.01.01 00:00:00",
            commission="-12", fee="0", swap="0", profit="0"),
    ])
    write_csv(later, [
        row("12", position="PX", entry="out", time="2026.02.01 00:00:00",
            symbol="LATER", commission="0", fee="0", swap="0", profit="2"),
    ])
    write_csv(earlier, [
        row("11", position="PX", entry="out", time="2026.01.15 00:00:00",
            symbol="EARLIER", commission="0", fee="0", swap="0", profit="1"),
    ])
    results = []
    for order in ([opening, later, earlier], [earlier, later, opening]):
        out = tmp / f"order-{len(results)}.xlsx"
        proses(order, out)
        wb = load_workbook(out, read_only=True, data_only=True)
        try:
            by_symbol = {r[3]: r[6] for r in values(wb["Daily"])[1:]}
            results.append(by_symbol)
        finally:
            wb.close()
    assert results == [
        {"EARLIER": -12, "LATER": 0},
        {"EARLIER": -12, "LATER": 0},
    ], results

# Untrusted text must stay literal in Excel; typed dates/numbers must remain typed.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "formula.csv", tmp / "formula.xlsx"
    write_csv(src, [row("formula", login="=1+1", country="+MY",
                        group=r"real\-Desk", type_="@buy", symbol="=EURUSD",
                        currency="+USD")])
    proses([src], out)
    wb = load_workbook(out, data_only=False)
    try:
        for sheet in ("Daily", "Monthly Summary"):
            cells = list(wb[sheet][2])
            assert [cells[i].value for i in (0, 2, 3, 10)] == [
                "=1+1", "@buy", "=EURUSD", "+USD",
            ]
            # quotePrefix keeps the value literal in Excel WITHOUT a stray apostrophe.
            assert all(cells[i].quotePrefix for i in (0, 2, 3, 10))
            assert cells[1].is_date
            assert all(c.data_type == "n" for c in cells[4:10])
    finally:
        wb.close()

# An unreadable XLSX reports the failure in English.
with TemporaryDirectory() as tmp:
    empty = Path(tmp) / "empty.xlsx"
    Workbook().save(empty)
    try:
        proses([empty], Path(tmp) / "out.xlsx")
    except SystemExit as exc:
        assert "no header row" in str(exc), exc
    else:
        raise AssertionError("empty workbook was accepted")

# Trust-boundary diagnostics include source location and offending column.
expect_error([row("1", volume="not-a-number")], "STOP:", "bad.csv", "row 2", "Volume")
expect_error([row("1", fee="")], "STOP:", "bad.csv", "row 2", "Fee")
expect_error([row("1", profit="NaN")], "STOP:", "bad.csv", "row 2", "Profit", "NaN")
expect_error([row("1", swap="inf")], "STOP:", "bad.csv", "row 2", "Swap", "inf")
expect_error([row("1", time="")], "STOP:", "bad.csv", "row 2", "Time")
expect_error([row("1", time="2026.02.30 00:00:00")], "STOP:", "bad.csv", "row 2", "Time")
expect_error([row("1", time="2026.07.01 25:00:00")], "STOP:", "bad.csv", "row 2", "Time")
expect_error([row("1", time="2026.07.01 garbage")], "STOP:", "bad.csv", "row 2", "Time")
expect_error([row("1", time="yesterday")], "STOP:", "bad.csv", "row 2", "Time")
expect_error([["1", "100"]], "STOP:", "bad.csv", "row 2", "expected 15 columns", "found 2 columns")

def expect_rejected_xlsx_closed(filename, populate, *message_parts):
    with TemporaryDirectory() as tmp:
        src = Path(tmp) / filename
        workbook = Workbook()
        populate(workbook.active)
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
            except SystemExit as exc:
                message = str(exc)
                for part in message_parts:
                    assert part in message, (part, message)
            else:
                raise AssertionError("invalid XLSX was accepted")
            assert opened and opened[0]._archive.fp is None
        finally:
            deal_segregator.load_workbook = real_load_workbook


# Rejected XLSX workbooks close during both header-validation and empty initialization.
expect_rejected_xlsx_closed("missing-columns.xlsx", lambda ws: ws.append(["Deal"]),
                            "STOP:", "missing-columns.xlsx", "required columns")
expect_rejected_xlsx_closed("empty.xlsx", lambda ws: None,
                            "STOP:", "empty.xlsx", "header")

# Used normalized rows must be released while input is still consumed.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "many.csv"
    write_csv(src, [row(str(i)) for i in range(200)])
    real_angka = deal_segregator._angka
    alive = peak = 0
    class TrackedFloat(float):
        def __new__(cls, value):
            nonlocal_alive[0] += 1
            nonlocal_alive[1] = max(nonlocal_alive[1], nonlocal_alive[0])
            return super().__new__(cls, value)
        def __del__(self):
            nonlocal_alive[0] -= 1
    nonlocal_alive = [alive, peak]
    deal_segregator._angka = lambda *args: TrackedFloat(real_angka(*args))
    try:
        summary = proses([src], tmp / "out.xlsx")
    finally:
        deal_segregator._angka = real_angka
    assert summary["dipakai"] == 200
    assert nonlocal_alive[1] < 20, f"normalized rows retained: peak={nonlocal_alive[1]}"

print("OK deal segregator engine")
