"""MT4 Raw Report support for Deal Segregator. python3 test_deal_segregator_mt4.py"""
import base64
import csv
import datetime
import io
import os
import tempfile
import time
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(ROOT / "users.db")
os.environ["DW_JOB_DIR"] = str(ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "mt4-tester"
os.environ["DW_ADMIN_PASS"] = "mt4-tester-password-1"

from deal_segregator import proses  # noqa: E402

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "kvb"
MT5_FIXTURE = FIXTURE / "KVB_MT5_Deals_2026-10-04.csv"
MT4_FIXTURE = FIXTURE / "KVB_MT4_RawReport_2026-10-04.csv"

MT5_HEADERS = [
    "Deal", "Login", "Time", "Type", "Entry", "Symbol", "Volume",
    "Commission", "Fee", "Swap", "Profit", "Currency",
]
MT4_HEADER = [
    "Deal", "Login", "Open Time", "Type", "Symbol", "Volume", "Open Price",
    "S/L", "T/P", "Close Time", "Close Price", "Reason", "Gateway Order",
    "Gateway Volume", "Open Price Delta", "Close Price Delta", "Agent",
    "Commission", "Taxes", "Swap", "Profit", "Points", "Comment",
]
MT4_RESULT = [
    "Login", "Date", "Type", "Symbol", "Deals", "Volume",
    "Commission", "Agent", "Fee", "Swap", "Profit", "Currency",
]

K_BAL = "MT4 balance/credit rows dropped (same rule as MT5 non-trade Entry)"
K_CAN = "MT4 cancelled pending orders dropped"
K_FOOT = "MT4 footer/summary lines dropped"
K_OPEN_OUT = "MT4 opening legs outside the report period (in row skipped)"
K_OUT = "MT4 closed buy/sell kept"
K_IN = "MT4 opening legs kept (volume only, inside the report period)"
K_DUP = "MT4 duplicate Deal IDs dropped"
K_OTHER = "MT4 other rows dropped"
K_CUR = "MT4 currency assumption (open question)"
K_QUEST = "MT4 open questions for KVB"
K_PROFIT = "MT4 Total Profit (closed buy/sell)"
K_AGENT_SUM = "MT4 Total Agent (closed buy/sell)"
K_FEE = "MT4 Total Fee (Taxes, closed buy/sell)"
K_DAILY = "Daily MT4 rows (Login+Date+Type+Symbol+Currency)"
K_DAILY_IN = "Daily MT4 in rows (sheet 'Daily - MT4 In')"


def write_mt5(path, rows):
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(MT5_HEADERS)
        w.writerows(rows)


def mt5_row(deal="1", login="100", time="2026.10.04 01:02:03", type_="buy",
            entry="out", symbol="EURUSD", volume="1", commission="0", fee="0",
            swap="0", profit="10", currency="USD"):
    return [deal, login, time, type_, entry, symbol, volume, commission, fee,
            swap, profit, currency]


def write_mt4(path, rows, start="2026.10.04", end="2026.10.04", footer=True):
    """Content is what identifies the file. The name is chosen by the caller."""
    with path.open("w", encoding="utf-8", newline="") as f:
        f.write(f"Raw Report for 'abcc' from {start} to {end}\n")
        w = csv.writer(f, delimiter=";")
        w.writerow(MT4_HEADER)
        for row in rows:
            w.writerow([row.get(h, "") for h in MT4_HEADER])
        if footer:
            w.writerow([""] * 16 + ["0.00", "-1.00", "0.00", "0.00", "10.00", "", ""])
            w.writerow([""] * 18 + ["Profit:", "", "10.00", "", ""])
            w.writerow([""] * 18 + ["Deposit:", "", "1.00", "", ""])
            w.writerow([""] * 18 + ["Withdrawal:", "", "-1.00", "", ""])
            w.writerow([""] * 18 + ["Credit:", "", "0.00", "", ""])


def mt4_trade(deal, login="100", open_time="2026.10.04 01:00:00", type_="buy",
              symbol="eurusd", volume="1", close_time="2026.10.04 02:00:00",
              agent="0", commission="0", taxes="0", swap="0", profit="0",
              comment=""):
    return {
        "Deal": deal, "Login": login, "Open Time": open_time, "Type": type_,
        "Symbol": symbol, "Volume": volume, "Close Time": close_time,
        "Agent": agent, "Commission": commission, "Taxes": taxes, "Swap": swap,
        "Profit": profit, "Comment": comment,
    }


def books(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        assert z.namelist() == ["Deals - Daily.xlsx", "Deals - Monthly Summary.xlsx"]
        for name in z.namelist():
            wb = load_workbook(io.BytesIO(z.read(name)), data_only=True)
            out[name] = {s: [tuple(r) for r in wb[s].iter_rows(values_only=True)]
                         for s in wb.sheetnames}
            out[name]["__sheets__"] = list(wb.sheetnames)
            wb.close()
    return out


def verification(rows):
    return {r[0]: r[1] for r in rows if r and r[0]}


def run(paths, allow_mt4=False):
    folder = Path(tempfile.mkdtemp())
    out = folder / "hasil.zip"
    summary = proses([Path(p) for p in paths], out, allow_mt4=allow_mt4)
    return summary, books(out), out


# The attached MT5 export is UTF-16, tab-separated, and ends with a Total row
# plus per-currency summary rows (including USC). Entry is blank on those
# rows, so the existing parser drops them and keeps USC as a currency.
assert MT5_FIXTURE.is_file() and MT4_FIXTURE.is_file()
summary, hasil, _ = run([MT5_FIXTURE], allow_mt4=False)
daily = hasil["Deals - Daily.xlsx"]
assert daily["__sheets__"] == ["Daily", "Daily - In", "Verifikasi"]
assert "Daily - MT4" not in daily
currencies = {r[10] for r in daily["Daily"][1:]}
assert "USC" in currencies and "USD" in currencies
v = verification(daily["Verifikasi"])
assert v["Total source rows (all files)"] == 7
assert v["Entry empty/other (dropped)"] == 3
assert v["Unique 'out' rows kept"] == 2
assert v["Unique 'in' rows kept"] == 2
assert K_OUT not in v and K_CUR not in v
assert summary["dipakai"] == 2 and summary["dipakai_in"] == 2
assert "mt4_dipakai" not in summary
print("OK MT5 fixture already parsed (USC kept, footer dropped)")

# Same MT5 file, KVB flag on, must not grow sheets or change verification text.
_, dengan_flag, _ = run([MT5_FIXTURE], allow_mt4=True)
assert dengan_flag["Deals - Daily.xlsx"]["Verifikasi"] == daily["Verifikasi"]
assert dengan_flag["Deals - Daily.xlsx"]["__sheets__"] == daily["__sheets__"]
assert dengan_flag["Deals - Daily.xlsx"]["Daily"] == daily["Daily"]
assert dengan_flag["Deals - Daily.xlsx"]["Daily - In"] == daily["Daily - In"]
assert dengan_flag["Deals - Monthly Summary.xlsx"]["__sheets__"] == [
    "Monthly Summary", "Monthly Summary - In", "Verifikasi"]
print("OK MT5-only output unchanged when MT4 is allowed")

# DPM default refuses MT4 by content, including when the name looks like MT5.
with tempfile.TemporaryDirectory() as tmp:
    mt4 = Path(tmp) / "Deals History.csv"
    mt5 = Path(tmp) / "Raw Report.csv"
    write_mt4(mt4, [mt4_trade("9")])
    write_mt5(mt5, [mt5_row()])
    try:
        proses([mt4, mt5], Path(tmp) / "out.zip")
    except SystemExit as exc:
        message = str(exc)
        assert "STOP:" in message and "MT5 Deals History only" in message, message
        assert "Deals History.csv" in message and "KVB" in message, message
    else:
        raise AssertionError("DPM accepted an MT4 Raw Report")
    summary, hasil, _ = run([mt5], allow_mt4=False)
    assert hasil["Deals - Daily.xlsx"]["__sheets__"] == ["Daily", "Daily - In", "Verifikasi"]
    assert hasil["Deals - Daily.xlsx"]["Daily"][1][0] == "100"
print("OK DPM rejects MT4 by content and still reads a misnamed MT5 file")

# Closed buy/sell -> out at Close Time. Taxes land in Fee. Agent stays its own
# column. Same Login+Date+Type+Symbol aggregates. An in row is volume-only and
# only when Open Time falls inside the report period.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "notes.csv"
    write_mt4(src, [
        mt4_trade("1", login="100", open_time="2026.10.03 23:00:00",
                  close_time="2026.10.04 01:00:00", symbol="btcusd", volume="0.02",
                  commission="-0.17", taxes="-1.5", swap="-0.25", profit="3.98",
                  agent="2"),
        mt4_trade("2", login="200", open_time="2026.10.04 03:00:00",
                  close_time="2026.10.04 04:00:00", type_="sell", symbol="ethusd",
                  volume="0.5", commission="-0.4", taxes="0", swap="0", profit="-8",
                  agent="0.5"),
        mt4_trade("3", login="100", open_time="2026.10.01 00:00:00",
                  close_time="2026.10.04 05:00:00", symbol="btcusd", volume="1",
                  profit="1", agent="0"),
        {"Deal": "4", "Login": "300", "Open Time": "2026.10.04 00:05:00",
         "Type": "balance", "Symbol": "CRM-S-1,Swap", "Profit": "-2.35",
         "Comment": "CRM-S-1,Swap"},
        {"Deal": "5", "Login": "301", "Open Time": "2026.10.04 00:06:00",
         "Type": "credit", "Profit": "9", "Comment": "credit in"},
        {"Deal": "6", "Login": "302", "Open Time": "2026.10.03 01:00:00",
         "Type": "buy limit", "Symbol": "btcusd", "Volume": "0.2",
         "Close Time": "2026.10.04 02:00:00", "Comment": "cancelled"},
        {"Deal": "7", "Login": "303", "Open Time": "2026.10.04 01:00:00",
         "Type": "sell stop", "Symbol": "ethusd", "Volume": "0.1",
         "Close Time": "2026.10.04 02:00:00", "Comment": "cancelled"},
    ])
    summary, hasil, _ = run([src], allow_mt4=True)
    daily = hasil["Deals - Daily.xlsx"]
    assert daily["__sheets__"] == ["Daily - MT4", "Daily - MT4 In", "Verifikasi"]
    monthly = hasil["Deals - Monthly Summary.xlsx"]
    assert monthly["__sheets__"] == [
        "Monthly Summary - MT4", "Monthly Summary - MT4 In", "Verifikasi"]
    assert len("Monthly Summary - MT4 In") <= 31
    assert daily["Daily - MT4"][0] == tuple(MT4_RESULT)
    assert daily["Daily - MT4"][1:] == [
        ("100", datetime.datetime(2026, 10, 4), "buy", "BTCUSD", 2, 1.02, -0.17, 2, -1.5, -0.25, 4.98, "USD"),
        ("200", datetime.datetime(2026, 10, 4), "sell", "ETHUSD", 1, 0.5, -0.4, 0.5, 0, 0, -8, "USD"),
    ]
    assert daily["Daily - MT4 In"][1:] == [
        ("200", datetime.datetime(2026, 10, 4), "sell", "ETHUSD", 1, 0.5, 0, 0, 0, 0, 0, "USD"),
    ]
    v = verification(daily["Verifikasi"])
    assert v[K_OUT] == 3 and v[K_IN] == 1
    assert v[K_BAL] == 2 and v[K_CAN] == 2 and v[K_FOOT] == 5
    assert v[K_OPEN_OUT] == 2 and v[K_DUP] == 0 and v[K_OTHER] == 0
    assert v[K_PROFIT] == round(3.98 + 1 - 8, 2)
    assert v[K_AGENT_SUM] == 2.5
    assert v[K_FEE] == -1.5
    assert v[K_DAILY] == 2 and v[K_DAILY_IN] == 1
    assert str(v[K_CUR]).startswith("USD")
    assert "open question" in K_CUR and "Agent" in v[K_QUEST] and "balance" in v[K_QUEST].lower()
    assert "Commission" in v["MT4 Agent commission"]
    assert v["MT4 balance/credit Profit excluded"] == round(-2.35 + 9, 2)
    assert summary["mt4_dipakai"] == 3 and summary["mt4_dipakai_in"] == 1
    # Money on the in row stays zero, so it is not added to the close.
    assert v["MT4 Total Commission (closed buy/sell)"] == round(-0.17 - 0.4, 2)
    file_line = next(val for key, val in v.items() if "notes.csv" in key and "MT4" in key)
    assert "balance/credit" in file_line and "cancelled" in file_line
print("OK MT4 parsing, footer, cancelled orders, balance rows, in-row rule")

# Period boundaries are inclusive. 1 Oct is inside a 1–4 Oct report; 30 Sep is not.
with tempfile.TemporaryDirectory() as tmp:
    src = Path(tmp) / "week.csv"
    write_mt4(src, [
        mt4_trade("a", login="1", open_time="2026.10.01 00:00:00",
                  close_time="2026.10.04 12:00:00", volume="1", profit="5"),
        mt4_trade("b", login="2", open_time="2026.09.30 23:59:59",
                  close_time="2026.10.04 12:00:00", volume="1", profit="7"),
    ], start="2026.10.01", end="2026.10.04", footer=False)
    _, hasil, _ = run([src], allow_mt4=True)
    daily = hasil["Deals - Daily.xlsx"]
    assert [r[0] for r in daily["Daily - MT4 In"][1:]] == ["1"]
    assert verification(daily["Verifikasi"])[K_OPEN_OUT] == 1
print("OK report-period boundaries")

# Duplicate Deal IDs across two MT4 files count once. Blank Login on a trade stops.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    first, second = tmp / "a.csv", tmp / "b.csv"
    write_mt4(first, [mt4_trade("same", login="5", profit="4", volume="1")], footer=False)
    write_mt4(second, [mt4_trade("same", login="5", profit="99", volume="9")], footer=False)
    _, hasil, _ = run([first, second], allow_mt4=True)
    v = verification(hasil["Deals - Daily.xlsx"]["Verifikasi"])
    assert v[K_DUP] == 1 and v[K_OUT] == 1
    assert hasil["Deals - Daily.xlsx"]["Daily - MT4"][1][10] == 4  # Profit, not 99
    blank = tmp / "blank.csv"
    write_mt4(blank, [mt4_trade("z", login="", profit="1")], footer=False)
    try:
        proses([blank], tmp / "out.zip", allow_mt4=True)
    except SystemExit as exc:
        assert "blank Login" in str(exc) and "blank.csv" in str(exc)
    else:
        raise AssertionError("blank MT4 Login was accepted")
print("OK MT4 duplicate deals and blank Login")

# MT4 and MT5 logins and Deal IDs can overlap. They stay on separate sheets.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    mt5, mt4 = tmp / "mt5.csv", tmp / "mt4.csv"
    write_mt5(mt5, [mt5_row(deal="42", login="100", symbol="EURUSD", profit="10")])
    write_mt4(mt4, [mt4_trade("42", login="100", symbol="eurusd", profit="3", volume="2",
                              agent="1.25", commission="-0.5")], footer=False)
    summary, hasil, _ = run([mt5, mt4], allow_mt4=True)
    daily = hasil["Deals - Daily.xlsx"]
    assert daily["__sheets__"] == [
        "Daily", "Daily - In", "Daily - MT4", "Daily - MT4 In", "Verifikasi"]
    assert daily["Daily"][1][9] == 10 and daily["Daily"][1][3] == "EURUSD"
    assert "Agent" not in daily["Daily"][0]
    mt4_row = daily["Daily - MT4"][1]
    assert mt4_row[0] == "100" and mt4_row[10] == 3 and mt4_row[7] == 1.25
    assert mt4_row[6] == -0.5  # Commission is not Commission+Agent
    v = verification(daily["Verifikasi"])
    assert v["Unique logins"] == 1 and v["Unique MT4 logins"] == 1
    assert v["Total Profit ('out' rows)"] == 10
    assert v[K_PROFIT] == 3
    assert summary["dipakai"] == 1 and summary["mt4_dipakai"] == 1
print("OK mixed MT4+MT5 do not merge overlapping logins or Deal IDs")

# FX converts MT4 money, including Agent, with the row's own close date.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    deals, fx = tmp / "mt4.csv", tmp / "fx.xlsx"
    write_mt4(deals, [mt4_trade(
        "1", login="8", open_time="2026.10.04 01:00:00",
        close_time="2026.10.04 09:00:00", profit="10", agent="4",
        commission="-2", taxes="-1", swap="-0.5", volume="1")], footer=False)
    wb = Workbook()
    ws = wb.active
    ws.title = "Query result"
    ws.append(["date", "Currency", "rate"])
    ws.append([datetime.date(2026, 10, 4), "USD", 2])
    wb.save(fx)
    wb.close()
    _, hasil, out = run([deals, fx], allow_mt4=True)
    wb = load_workbook(io.BytesIO(zipfile.ZipFile(out).read("Deals - Daily.xlsx")))
    try:
        headers = [c.value for c in wb["Daily - MT4"][1]]
        assert "Agent (USD)" in headers and "Commission (USD)" in headers
        row = {headers[i]: c.value for i, c in enumerate(wb["Daily - MT4"][2])}
        assert row["Profit (USD)"] == 5
        assert row["Agent (USD)"] == 2
        assert row["Commission (USD)"] == -1
        assert row["Fee (USD)"] == -0.5
        assert row["Swap (USD)"] == -0.25
        in_headers = [c.value for c in wb["Daily - MT4 In"][1]]
        in_row = {in_headers[i]: c.value for i, c in enumerate(wb["Daily - MT4 In"][2])}
        assert in_row["Agent"] == 0 and in_row["Agent (USD)"] == 0
    finally:
        wb.close()
print("OK MT4 FX includes Agent")

# Trimmed copies of the attached 2026-10-04 samples.
summary, hasil, _ = run([MT5_FIXTURE, MT4_FIXTURE], allow_mt4=True)
daily = hasil["Deals - Daily.xlsx"]
assert daily["__sheets__"] == [
    "Daily", "Daily - In", "Daily - MT4", "Daily - MT4 In", "Verifikasi"]
symbols = {r[3] for r in daily["Daily - MT4"][1:]}
assert symbols == {"BTCUSD"}
logins = {r[0] for r in daily["Daily - MT4"][1:]}
assert "603477" in logins and "20033032" not in logins and "20023282" not in logins
# Opened the day before the report: out only.
opened_earlier = [r for r in daily["Daily - MT4"][1:] if r[0] == "603477"]
assert opened_earlier and opened_earlier[0][10] == 3.98
assert all(r[0] != "603477" for r in daily["Daily - MT4 In"][1:])
# Same-day open is on the in sheet, volume only.
same_day = [r for r in daily["Daily - MT4 In"][1:] if r[0] == "20038203"]
assert same_day == [("20038203", datetime.datetime(2026, 10, 4), "buy", "BTCUSD",
                     1, 0.01, 0, 0, 0, 0, 0, "USD")]
v = verification(daily["Verifikasi"])
assert v[K_BAL] == 2 and v[K_CAN] == 2 and v[K_FOOT] == 7
assert v[K_OPEN_OUT] == 1 and v[K_OUT] == 3 and v[K_IN] == 2
assert v["Entry empty/other (dropped)"] == 3  # MT5 footer, unchanged
assert "USC" in {r[10] for r in daily["Daily"][1:]}
assert summary["dipakai"] == 2 and summary["mt4_dipakai"] == 3
print("OK trimmed KVB MT4+MT5 fixture")

# Formula-like Login stays literal on the MT4 sheet.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "formula.csv"
    write_mt4(src, [mt4_trade("1", login="=1+1", symbol="=eurusd", profit="1")], footer=False)
    _, _, out = run([src], allow_mt4=True)
    wb = load_workbook(io.BytesIO(zipfile.ZipFile(out).read("Deals - Daily.xlsx")))
    try:
        login, _date, _type, symbol = list(wb["Daily - MT4"][2])[:4]
        assert login.value == "=1+1" and symbol.value == "=EURUSD"
        assert login.quotePrefix and login.data_type == "s"
        assert symbol.quotePrefix and symbol.data_type == "s"
    finally:
        wb.close()
print("OK MT4 formula text stays literal")


def wait(client, location):
    job_id = location.rstrip("/").rsplit("/", 1)[-1]
    for _ in range(200):
        state = __import__("app")._read_state(job_id)
        if state and state["state"] in ("done", "failed"):
            return job_id, state
        time.sleep(0.02)
    raise AssertionError("job timeout")


import app as A  # noqa: E402

A.app.config["TESTING"] = True
token = base64.b64encode(b"mt4-tester:mt4-tester-password-1").decode()
client = A.app.test_client()
client.environ_base["HTTP_AUTHORIZATION"] = f"Basic {token}"
client.environ_base["HTTP_ACCEPT"] = "text/html"

kvb_page = client.get("/kvb/tool/segregate")
assert kvb_page.status_code == 200
assert b"MT4 Raw Report is optional" in kvb_page.data
dpm_page = client.get("/tool/segregate")
assert dpm_page.status_code == 200
assert b"MT4 Raw Report is optional" not in dpm_page.data
assert b"Drop Deals History files here" in dpm_page.data
kvb_home = client.get("/kvb")
assert b"MT4 Raw Report is optional" in kvb_home.data
print("OK KVB tool page mentions optional MT4; DPM page does not")

mt4_body = (
    "Raw Report for 'abcc' from 2026.10.04 to 2026.10.04\n"
    + ";".join(MT4_HEADER) + "\n"
    + ";".join(["9", "100", "2026.10.04 01:00:00", "buy", "btcusd", "1", "", "", "",
                "2026.10.04 02:00:00", "", "", "", "", "", "", "0", "-1", "0", "0", "4", "", ""])
    + "\n"
).encode()

dpm = client.post("/tool/segregate", data={"file": (io.BytesIO(mt4_body), "Deals History.csv")},
                  content_type="multipart/form-data")
assert dpm.status_code == 303, (dpm.status_code, dpm.data[:300])
_, dpm_state = wait(client, dpm.headers["Location"])
assert dpm_state["state"] == "failed", dpm_state
assert "MT5 Deals History only" in (dpm_state.get("error") or "")
assert "KVB" in (dpm_state.get("error") or "")

commands = []
real_run = A.subprocess.run


def capture(command, **kwargs):
    commands.append(command)
    return real_run(command, **kwargs)


A.subprocess.run = capture
try:
    kvb = client.post("/kvb/tool/segregate",
                      data={"file": (io.BytesIO(mt4_body), "whatever.csv")},
                      content_type="multipart/form-data")
    assert kvb.status_code == 303, (kvb.status_code, kvb.data[:300])
    _, kvb_state = wait(client, kvb.headers["Location"])
    assert kvb_state["state"] == "done", kvb_state
    assert commands and "--allow-mt4" in commands[-1]
    assert Path(commands[-1][1]).name == "deal_segregator.py"
finally:
    A.subprocess.run = real_run
print("OK KVB runs MT4 and DPM's job rejects it")

print("OK deal segregator MT4")
