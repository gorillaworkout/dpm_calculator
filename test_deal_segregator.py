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
            deal_segregator._tulis_zip([row], [row], [], [], out, {}, False, set())
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
        assert wb.sheetnames == ["Daily", "Daily - In", "Monthly Summary",
                                 "Monthly Summary - In", "Verifikasi"]
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
        assert verification["Entry = in (before deduplication)"] == 1
        assert verification["Duplicate non-empty Deal IDs (dropped)"] == 1
        assert verification["Unique 'out' rows kept"] == 5
        assert verification["Unique 'in' rows kept"] == 1
        # The 'in' row lives on its own sheet, with its own values as in MT5.
        assert values(wb["Daily - In"])[0] == tuple(RESULT_HEADERS)
        assert values(wb["Daily - In"])[1:] == [
            ("100", datetime.datetime(2026, 7, 1), "buy", "EURUSD", 1, 1.25, -1, -0.1, -0.2, 999, "USD"),
        ]
        assert len(values(wb["Monthly Summary - In"])) == 2
        # The label must state every grouping dimension actually used.
        assert verification["Daily 'out' rows (Login+Date+Type+Symbol+Currency)"] == 4
        assert verification["Monthly Summary 'out' rows (Login+Month+Type+Symbol+Currency)"] == 3
        assert verification["Unique logins"] == 2
        assert "Rows with Country" not in verification
        assert "Rows without Country" not in verification
        assert "Unique desks" not in verification
        assert verification["Total Profit ('out' rows)"] == 72
        assert verification["Total Profit ('in' rows)"] == 999
        assert summary == {"file_masuk": 2, "total": 7, "dipakai": 5, "dipakai_in": 1,
                           "harian": 4, "bulanan": 3, "harian_in": 1, "bulanan_in": 1,
                           "login_unik": 2}
    finally:
        wb.close()

# 'in' rows get their OWN sheet (dikonfirmasi user 6 Okt 2026, opsi A): nothing is
# merged onto the 'out' rows, so Commission/Fee/Swap of an opening leg stays on the
# 'in' sheet exactly as MT5 reports it and is never counted twice.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "split.csv", tmp / "split.xlsx"
    write_csv(src, [
        row("1", position="P1", entry="in", type_="buy", commission="-2", fee="0", swap="0", profit="0"),
        row("2", position="P1", entry="out", type_="sell", commission="0", fee="0", swap="0", profit="10"),
        row("3", position="P2", entry="in", type_="buy", symbol="XAUUSD", commission="-5", fee="-0.5", swap="0", profit="0"),
        row("1", position="P1", entry="in", type_="buy", commission="-2", fee="0", swap="0", profit="0"),  # duplicate Deal
    ])
    summary = proses([src], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        out_rows = values(wb["Daily"])[1:]
        in_rows = {r[3]: r for r in values(wb["Daily - In"])[1:]}
        assert [r[3:] for r in out_rows] == [("EURUSD", 1, 1.25, 0, 0, 0, 10, "USD")], \
            "'out' sheet must show the 'out' row exactly as MT5 (no merged commission)"
        assert in_rows["EURUSD"][4:10] == (1, 1.25, -2, 0, 0, 0)
        assert in_rows["XAUUSD"][6] == -5 and in_rows["XAUUSD"][7] == -0.5
        verification = dict(values(wb["Verifikasi"])[2:])
        assert verification["Total Commission ('out' rows)"] == 0
        assert verification["Total Commission ('in' rows)"] == -7
        assert verification["Duplicate non-empty Deal IDs (dropped)"] == 1
        assert summary["dipakai"] == 1 and summary["dipakai_in"] == 2
    finally:
        wb.close()

# A multi-sheet XLSX (sheet IN + sheet OUT, like the Malaysia team's check file) must be
# read in full: every sheet with Deals History columns, columns matched by NAME, other
# sheets skipped and reported.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "two-sheets.xlsx", tmp / "two-sheets-out.xlsx"
    workbook = Workbook()
    ws_in, ws_out = workbook.active, workbook.create_sheet("OUT")
    ws_in.title = "IN"
    ws_in.append(HEADERS)
    ws_in.append(row("1", entry="in", commission="-3", profit="0"))
    swapped = ["Login", "Deal"] + [h for h in HEADERS if h not in ("Login", "Deal")]  # other column order
    ws_out.append(swapped)
    r = dict(zip(HEADERS, row("2", entry="out", type_="sell", commission="0", profit="8")))
    ws_out.append([r[h] for h in swapped])
    workbook.create_sheet("Notes").append(["just", "a", "note"])
    workbook.save(src)
    workbook.close()
    summary = proses([src], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        assert values(wb["Daily - In"])[1][6] == -3
        assert values(wb["Daily"])[1][9] == 8
        verification = dict(values(wb["Verifikasi"])[2:])
        assert verification["      sheet 'IN'"] == "1 rows read"
        assert verification["      sheet 'OUT'"] == "1 rows read"
        assert any("skipped" in k and v == "Notes" for k, v in verification.items()), verification
        assert summary["dipakai"] == 1 and summary["dipakai_in"] == 1
    finally:
        wb.close()

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
            assert all(cells[i].data_type == "s" for i in (0, 2, 3, 10))
            assert cells[1].is_date
            assert all(c.data_type == "n" for c in cells[4:10])
        with zipfile.ZipFile(out) as archive:
            xml = archive.read("xl/worksheets/sheet1.xml")
            assert b"<f>1+1</f>" not in xml
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

# Position is not a required column, and a blank Position is kept as its own
# row. Opening charges stay on the in sheet; they are not rejected for lacking
# a Position to merge onto.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "no-position.csv", tmp / "no-position.xlsx"
    headers = [h for h in HEADERS if h != "Position"]
    sample = row("1", entry="out", commission="-3", profit="4")
    data = [c for h, c in zip(HEADERS, sample) if h != "Position"]
    with src.open("w", encoding="utf-8", newline="") as f:
        csv.writer(f).writerows([headers, data])
    proses([src], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        assert values(wb["Daily"])[1][6] == -3
        assert values(wb["Daily"])[1][9] == 4
        assert values(wb["Daily - In"])[1:] == []
    finally:
        wb.close()

with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "blank-position.csv", tmp / "blank-position.xlsx"
    write_csv(src, [
        row("in-blank", position="", entry="in", commission="-7", fee="-1", swap="-2", profit="0"),
        row("out-blank", position="", entry="out", commission="0", fee="0", swap="0", profit="10"),
    ])
    proses([src], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        assert values(wb["Daily"])[1][6:10] == (0, 0, 0, 10)
        assert values(wb["Daily - In"])[1][6:9] == (-7, -1, -2)
    finally:
        wb.close()

# Grouping identifiers are financial dimensions; blank Login would collapse
# unrelated unidentified accounts into one aggregate.
expect_error([row("blank-login", login="")], "STOP:", "blank Login", "row 2")
expect_error([
    row("blank-login-in", login="", position="PBL", entry="in",
        commission="-9", fee="-1", swap="-2", profit="0"),
    row("valid-close", login="100", position="PBL", entry="out",
        commission="0", fee="0", swap="0", profit="10"),
], "STOP:", "blank Login", "financial charges", "row 2")

# Volume is not money and must retain source precision.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "volume.csv", tmp / "volume.xlsx"
    write_csv(src, [row("small-volume", volume="0.001")])
    proses([src], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        assert values(wb["Daily"])[1][5] == 0.001
    finally:
        wb.close()

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

# OOM regression: high-cardinality Deal IDs must live on disk, not RAM.
# 150k unique closing rows + 150k opening rows must stay far below the ~3x footprint
# the old in-memory sets/dicts needed (production was OOM-killed at 7.2 GB).
import resource
import subprocess
import sys
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "big.csv", tmp / "big.zip"
    with src.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADERS)
        for i in range(150_000):
            w.writerow([f"in{i}", f"P{i}", str(100 + i % 50), "", "", "2026.07.01 00:00:00",
                        "buy", "in", "EURUSD", "1", "-1", "0", "0", "0", "USD"])
            w.writerow([f"out{i}", f"P{i}", str(100 + i % 50), "", "", "2026.07.02 00:00:00",
                        "buy", "out", "EURUSD", "1", "0", "0", "0", "1", "USD"])
    code = ("import resource,sys;from pathlib import Path;from deal_segregator import proses;"
            "r=proses([Path(sys.argv[1])],Path(sys.argv[2]));"
            "print(r['dipakai'],resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)")
    res = subprocess.run([sys.executable, "-c", code, str(src), str(out)], capture_output=True,
                         text=True, check=True, cwd=Path(__file__).parent)
    kept, rss = map(int, res.stdout.split())
    rss_mb = rss / (1024 * 1024 if sys.platform == "darwin" else 1024)
    assert kept == 150_000, kept
    assert rss_mb < 250, f"peak RSS {rss_mb:.0f} MB; high-cardinality state back in RAM"
print("OK deal segregator bounded memory")

# Sheet names stay put until a table passes the Excel row limit.
contoh = {"Login": "1", "Periode": datetime.date(2026, 9, 1), "Type": "buy",
          "Symbol": "EURUSD", "Deals": 1, "Volume": 1.0,
          "Commission": -1.0, "Fee": 0.0, "Swap": 0.0,
          "Profit": 2.0, "Currency": "USD", "_kurs_kurang": 0}
banyak = []
for i in range(5):
    item = dict(contoh)
    item["Login"] = str(i)
    banyak.append(item)
batas_asli = deal_segregator.EXCEL_DATA_MAKS
deal_segregator.EXCEL_DATA_MAKS = 2
try:
    with TemporaryDirectory() as folder:
        folder = Path(folder)
        out = folder / "split.zip"
        deal_segregator._tulis_zip(banyak, [], [], [], out, {}, False, set())
        with zipfile.ZipFile(out) as archive:
            archive.extract("Deals - Daily.xlsx", folder)
        wb = load_workbook(folder / "Deals - Daily.xlsx", read_only=True, data_only=True)
        try:
            assert wb.sheetnames == ["Daily", "Daily 2", "Daily 3", "Daily - In", "Verifikasi"]
            assert sum(1 for _ in wb["Daily"].iter_rows()) == 3
            assert sum(1 for _ in wb["Daily 3"].iter_rows()) == 2
            label = [r[0] for r in wb["Verifikasi"].iter_rows(values_only=True) if r]
            assert "Excel row limit" in label
        finally:
            wb.close()
finally:
    deal_segregator.EXCEL_DATA_MAKS = batas_asli

# The web app passes hapus_sumber. A direct call does not delete, and it does
# not claim that it did.
with TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, out = tmp / "keep.csv", tmp / "keep.zip"
    write_csv(src, [row("9")])
    proses([src], out)
    assert src.is_file()
    with zipfile.ZipFile(out) as archive:
        archive.extract("Deals - Daily.xlsx", tmp)
    wb = load_workbook(tmp / "Deals - Daily.xlsx", read_only=True, data_only=True)
    try:
        label = [r[0] for r in wb["Verifikasi"].iter_rows(values_only=True) if r]
    finally:
        wb.close()
    assert "Uploaded files on the server" not in label
    src2, out2 = tmp / "gone.csv", tmp / "gone.zip"
    write_csv(src2, [row("10")])
    proses([src2], out2, hapus_sumber=True)
    assert not src2.exists() and out2.is_file()
    (tmp / "extracted").mkdir()
    with zipfile.ZipFile(out2) as archive:
        archive.extract("Deals - Daily.xlsx", tmp / "extracted")
    wb = load_workbook(tmp / "extracted" / "Deals - Daily.xlsx", read_only=True, data_only=True)
    try:
        verifikasi = {r[0]: r[1] for r in wb["Verifikasi"].iter_rows(values_only=True)
                      if r and len(r) > 1 and r[0]}
    finally:
        wb.close()
    assert "Deleted from the server" in verifikasi["Uploaded files on the server"]
    assert "not changed" in verifikasi["Uploaded files on the server"]
print("OK excel split and delete-input")
