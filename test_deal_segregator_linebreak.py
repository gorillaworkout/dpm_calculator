#!/usr/bin/env python3
"""Line breaks inside a Deals History field must be repaired, not dropped.

python3 test_deal_segregator_linebreak.py
"""
import csv
import io
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

from deal_segregator import proses

ROOT = Path(__file__).resolve().parent
MT5 = ROOT / "fixtures" / "kvb" / "KVB_MT5_Deals_2026-10-04.csv"
MT4 = ROOT / "fixtures" / "kvb" / "KVB_MT4_Deals_2026-07-01_sample.csv"
LOGIN = "50076706"
NAME = "Sarayut Pansirichai"
K_REPAIR = "rows repaired (line break inside a field)"


def load_mt5():
    with MT5.open(encoding="utf-16", newline="") as handle:
        rows = list(csv.reader(handle, delimiter="\t"))
    return rows[0], rows[1:]


def write_mt5(path, header, data_rows):
    with path.open("w", encoding="utf-16", newline="") as handle:
        writer = csv.writer(handle, delimiter="\t")
        writer.writerow(header)
        writer.writerows(data_rows)


def write_mt5_raw(path, chunks):
    """chunks are physical lines already joined with tabs, without a newline."""
    text = "\n".join(chunks) + "\n"
    path.write_bytes(text.encode("utf-16"))


def verification(path):
    with zipfile.ZipFile(path) as archive:
        blob = archive.read("Deals - Daily.xlsx")
    book = load_workbook(io.BytesIO(blob), data_only=True)
    try:
        rows = list(book["Verifikasi"].iter_rows(values_only=True))
    finally:
        book.close()
    return {row[0]: row[1] for row in rows if row and row[0]}


def daily_deals(path):
    with zipfile.ZipFile(path) as archive:
        blob = archive.read("Deals - Daily.xlsx")
    book = load_workbook(io.BytesIO(blob), data_only=True)
    try:
        rows = list(book["Daily"].iter_rows(values_only=True))
    finally:
        book.close()
    header = rows[0]
    return [dict(zip(header, row)) for row in rows[1:]]


def run(path, allow_mt4=False):
    folder = Path(tempfile.mkdtemp())
    out = folder / "hasil.zip"
    proses([path], out, allow_mt4=allow_mt4)
    return verification(out), daily_deals(out), out


def repaired(summary):
    hits = [(key, value) for key, value in summary.items() if K_REPAIR in str(key)]
    assert hits, list(summary)
    return hits


def totals(summary, mt4=False):
    if mt4:
        return (
            summary["MT4 Total Profit (closed buy/sell)"],
            summary["MT4 Total Commission (closed buy/sell)"],
            summary["MT4 Total Volume (closed buy/sell)"],
            summary["MT4 closed buy/sell kept"],
        )
    return (
        summary["Total Profit ('out' rows)"],
        summary["Total Commission ('out' rows)"],
        summary["Total Volume ('out' rows)"],
        summary["Unique 'out' rows kept"],
    )


def expect_stop(path, login, allow_mt4=False):
    try:
        run(path, allow_mt4=allow_mt4)
    except SystemExit as exc:
        message = str(exc)
    else:
        raise AssertionError("a broken row was accepted")
    assert path.name in message, message
    assert str(login) in message, message
    assert "line" in message.lower() or "row" in message.lower(), message
    return message


header, data = load_mt5()
name_at = header.index("Name")
login_at = header.index("Login")
# Three consecutive deals, rewritten onto one login, then split.
base = [list(row) for row in data if row and row[0].isdigit()]
assert len(base) >= 3
clean_rows = [list(row) for row in data]
for row in clean_rows:
    if row and row[0].isdigit() and row in base[:3] or (row and row[0] in {base[0][0], base[1][0], base[2][0]}):
        pass
targets = {base[0][0], base[1][0], base[2][0]}
for row in clean_rows:
    if row and row[0] in targets:
        row[login_at] = LOGIN
        row[name_at] = NAME

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    clean = tmp / "clean.csv"
    write_mt5(clean, header, clean_rows)
    clean_summary, _, _ = run(clean)
    assert K_REPAIR not in "".join(map(str, clean_summary))
    print("OK clean MT5 file has no repair line")

    # Quoted newline in Name. csv.writer quotes it; the reader must keep the deal.
    quoted = tmp / "quoted.csv"
    quoted_rows = [list(row) for row in clean_rows]
    for row in quoted_rows:
        if row and row[login_at] == LOGIN:
            row[name_at] = "Sarayut\nPansirichai"
            break
    write_mt5(quoted, header, quoted_rows)
    summary, _, _ = run(quoted)
    assert totals(summary) == totals(clean_summary)
    text = " ".join(f"{key} {value}" for key, value in repaired(summary))
    assert LOGIN in text, text
    print("OK quoted newline in Name")

    # Unquoted newline: the first physical line stops inside Name.
    def unquoted_lines(rows, split_deals):
        lines = ["\t".join(header)]
        for row in rows:
            if row and row[0] in split_deals:
                left, right = row[name_at].split(" ", 1)
                lines.append("\t".join(row[:name_at] + [left]))
                lines.append("\t".join([right] + row[name_at + 1:]))
            else:
                lines.append("\t".join(row))
        return lines

    split_ids = [base[0][0], base[1][0], base[2][0]]
    raw = tmp / "unquoted.csv"
    write_mt5_raw(raw, unquoted_lines(clean_rows, set(split_ids)))
    summary, deals, _ = run(raw)
    assert totals(summary) == totals(clean_summary)
    text = " ".join(f"{key} {value}" for key, value in repaired(summary))
    assert LOGIN in text, text
    assert all(row["Login"] != LOGIN or row["Profit"] is not None for row in deals)
    print("OK unquoted newline in Name")

    # The same split is the last record in the file.
    tail = tmp / "unquoted-end.csv"
    body = [row for row in clean_rows if not (row and row[0] in split_ids)]
    last = [row for row in clean_rows if row and row[0] == split_ids[0]]
    write_mt5_raw(tail, unquoted_lines(body + last, {split_ids[0]}))
    write_mt5(tmp / "tail-clean.csv", header, body + last)
    assert totals(run(tail)[0]) == totals(run(tmp / "tail-clean.csv")[0])
    assert LOGIN in " ".join(str(v) for _, v in repaired(run(tail)[0]))
    print("OK unquoted newline at end of file")

    # The file ends on the first half of the name. That cannot be repaired.
    cut = tmp / "cut.csv"
    lines = unquoted_lines(body + last, {split_ids[0]})
    write_mt5_raw(cut, lines[:-1])
    message = expect_stop(cut, LOGIN)
    assert "line" in message.lower() or "row" in message.lower()
    print("OK file ending inside a field stops")

    # The next physical line makes the row wider than the header.
    wide = tmp / "wide.csv"
    bad_lines = ["\t".join(header), "\t".join(last[0][:name_at] + ["Sarayut"])]
    bad_lines.append("\t".join(["Pansirichai"] + last[0][name_at + 1:] + ["extra"]))
    write_mt5_raw(wide, bad_lines)
    message = expect_stop(wide, LOGIN)
    print("OK unrepairable CSV row stops")

    # Excel scientific notation is text, so the deal id has already lost digits.
    sci = tmp / "sci.csv"
    sci_rows = [list(row) for row in clean_rows]
    sci_rows[0][0] = "1.41E+08"
    write_mt5(sci, header, sci_rows)
    try:
        run(sci)
    except SystemExit as exc:
        message = str(exc)
    else:
        raise AssertionError("scientific notation was accepted")
    assert sci.name in message and "1.41E+08" in message, message
    assert "original" in message.lower(), message
    print("OK scientific-notation Deal stops")

    # Re-saved xlsx: several consecutive deals of one login are each two rows.
    def add_split(sheet, row):
        left, right = row[name_at].split(" ", 1)
        first = row[:name_at] + [left] + [None] * (len(header) - name_at - 1)
        second = [right] + row[name_at + 1:]
        sheet.append(first)
        sheet.append(second)

    book = Workbook()
    sheet = book.active
    sheet.append(header)
    seen = 0
    for row in clean_rows:
        if row and row[0] in split_ids:
            add_split(sheet, row)
            seen += 1
        else:
            sheet.append(row)
    assert seen == 3
    # A normal row sits between the split run and the summary, and one before it.
    xlsx = tmp / "split.xlsx"
    book.save(xlsx)
    book.close()
    summary, _, _ = run(xlsx)
    assert totals(summary) == totals(clean_summary)
    text = " ".join(f"{key} {value}" for key, value in repaired(summary))
    assert LOGIN in text and "3" in text, text
    print("OK xlsx split rows")

    # The following row cannot be joined into a deal.
    book = Workbook()
    sheet = book.active
    sheet.append(header)
    broken = list(last[0])
    left = "Sarayut"
    sheet.append(broken[:name_at] + [left] + [None] * (len(header) - name_at - 1))
    sheet.append(["not a date", "still not a deal"])
    bad_xlsx = tmp / "bad.xlsx"
    book.save(bad_xlsx)
    book.close()
    message = expect_stop(bad_xlsx, LOGIN)
    print("OK unrepairable xlsx row stops")

print("OK MT5 line breaks")

# MT4 comma export: an unquoted break in Symbol, which is not the last column.
with MT4.open(encoding="cp1252", newline="") as handle:
    mt4_rows = list(csv.reader(handle))
mt4_header = mt4_rows[0]
symbol_at = mt4_header.index("Symbol")
login_mt4 = mt4_header.index("Login")
trades = [row for row in mt4_rows[1:] if len(row) > 3 and row[3] in ("buy", "sell")]
assert trades
sample = [list(row) for row in trades[:5]]
sample[0][login_mt4] = LOGIN
held = sample[0][symbol_at]
sample[0][symbol_at] = "xau usd"

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    clean = tmp / "mt4-clean.csv"
    broken = tmp / "mt4-break.csv"
    with clean.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(mt4_header)
        writer.writerows(sample)
    with broken.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(mt4_header)
        for index, row in enumerate(sample):
            if index == 0:
                writer.writerow(row[:symbol_at] + ["xau"])
                writer.writerow(["usd"] + row[symbol_at + 1:])
            else:
                writer.writerow(row)
    clean_summary, _, _ = run(clean, allow_mt4=True)
    summary, _, _ = run(broken, allow_mt4=True)
    assert totals(summary, mt4=True) == totals(clean_summary, mt4=True)
    text = " ".join(f"{key} {value}" for key, value in repaired(summary))
    assert LOGIN in text, text
    assert held.replace(" ", "") != ""  # the original symbol was non-empty
    print("OK MT4 unquoted newline")

print("OK line breaks")
