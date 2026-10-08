"""MT4 account-type currency and the header-first layout.

python3 test_deal_segregator_mt4_currency.py
"""
import csv
import datetime
import io
import os
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(ROOT / "users.db")
os.environ["DW_JOB_DIR"] = str(ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "mt4-cur"
os.environ["DW_ADMIN_PASS"] = "mt4-cur-password-1"

from deal_segregator import proses  # noqa: E402

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "kvb"
JULY = FIXTURE / "KVB_MT4_Deals_2026-07-01_sample.csv"
MAPPING = FIXTURE / "KVB_MT4_Account_Types_sample.csv"
MT5 = FIXTURE / "KVB_MT5_Deals_2026-10-04.csv"

MT4_HEADER = [
    "Deal", "Login", "Open Time", "Type", "Symbol", "Volume", "Open Price",
    "S/L", "T/P", "Close Time", "Close Price", "Reason", "Gateway Order",
    "Gateway Volume", "Open Price Delta", "Close Price Delta", "Agent",
    "Commission", "Taxes", "Swap", "Profit", "Points", "Comment",
]
K_PROFIT = "MT4 Total Profit (closed buy/sell)"
K_RAW = "MT4 Total Profit raw (before USC ÷100)"
K_COMM = "MT4 Total Commission (closed buy/sell)"
K_COMM_RAW = "MT4 Total Commission raw (before USC ÷100)"
K_SWAP = "MT4 Total Swap (closed buy/sell)"
K_FEE = "MT4 Total Fee (Taxes, closed buy/sell)"
K_VOL = "MT4 Total Volume (closed buy/sell)"
K_AGENT = "MT4 Agent column (sum of kept rows; not included in Profit, Commission, or Fee)"
K_MAP_ROWS = "MT4 currency from account-type mapping (rows)"
K_MAP_LOGINS = "MT4 currency from account-type mapping (logins)"
K_COL_ROWS = "MT4 currency from the currency column (rows)"
K_COL_LOGINS = "MT4 currency from the currency column (logins)"
K_FB_ROWS = "MT4 currency from the per-file USD/USC choice (rows)"
K_FB_LOGINS = "MT4 currency from the per-file USD/USC choice (logins)"
K_FB_LIST = "MT4 logins still on the per-file currency fallback"
K_DISAGREE = "MT4 account type and currency column disagree"
K_USC = "MT4 USC ÷100"
K_BAL = "MT4 balance/credit rows dropped (same rule as MT5 non-trade Entry)"
K_CAN = "MT4 cancelled pending orders dropped"


def books(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            wb = load_workbook(io.BytesIO(z.read(name)), data_only=True)
            out[name] = {s: [tuple(r) for r in wb[s].iter_rows(values_only=True)]
                         for s in wb.sheetnames}
            wb.close()
    return out


def verification(rows):
    return {r[0]: r[1] for r in rows if r and r[0]}


def run(paths, allow_mt4=False, mt4_currency=None, mt4_accounts=None):
    folder = Path(tempfile.mkdtemp())
    out = folder / "hasil.zip"
    summary = proses([Path(p) for p in paths], out, allow_mt4=allow_mt4,
                     mt4_currency=mt4_currency, mt4_accounts=mt4_accounts)
    return summary, books(out), out


def write_comma(path, rows, currency=True):
    """Header-first comma layout. currency=True adds the blank trailing column."""
    header = list(MT4_HEADER) + (["", ""] if currency else [])
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        for row in rows:
            cells = [row.get(h, "") for h in MT4_HEADER]
            if currency:
                cells.extend(["", row.get("Currency", "")])
            w.writerow(cells)


def trade(deal, login="10", type_="buy", symbol="xauusd", volume="1",
          profit="100", commission="0", swap="0", taxes="0", agent="0",
          comment="", currency="USD", close_time="2026.07.01 01:00:00"):
    return {
        "Deal": deal, "Login": login, "Open Time": "2026.07.01 00:10:00",
        "Type": type_, "Symbol": symbol, "Volume": volume,
        "Close Time": close_time, "Agent": agent, "Commission": commission,
        "Taxes": taxes, "Swap": swap, "Profit": profit, "Comment": comment,
        "Currency": currency,
    }


def write_mapping(path, pairs):
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["account", "account_type"])
        w.writerows(pairs)


def daily_rows(hasil):
    return hasil["Deals - Daily.xlsx"]["Daily"][1:]


def by_login(hasil):
    rows = hasil["Deals - Daily.xlsx"]["Daily"]
    header = rows[0]
    return {r[header.index("Login")]: dict(zip(header, r)) for r in rows[1:]}


# Header-first comma file: column currency, #N/A is not a currency, volume stays.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "july.csv"
    write_comma(src, [
        trade("1", login="10", profit="100", commission="-2", swap="-1",
              taxes="-4", volume="0.5", agent="9", currency="USD"),
        trade("2", login="20", profit="250", commission="-10", volume="3",
              currency="USC"),
        trade("3", login="30", type_="sell limit", profit="5", currency="#N/A",
              comment="cancelled"),
        {"Deal": "4", "Login": "40", "Open Time": "2026.07.01 00:01:00",
         "Type": "balance", "Symbol": "CRM-D-1", "Profit": "13.99",
         "Comment": "CRM-D-1", "Currency": "#N/A"},
    ])
    _, hasil, _ = run([src], allow_mt4=True)
    got = by_login(hasil)
    assert set(got) == {"10", "20"}
    assert got["10"]["Currency"] == "USD"
    assert got["10"]["Profit"] == 100 and got["10"]["Commission"] == -2
    assert got["10"]["Volume"] == 0.5
    assert got["20"]["Currency"] == "USC"
    assert got["20"]["Profit"] == 2.5 and got["20"]["Commission"] == -0.1
    assert got["20"]["Volume"] == 3
    assert "Agent" not in hasil["Deals - Daily.xlsx"]["Daily"][0]
    v = verification(hasil["Deals - Daily.xlsx"]["Verifikasi"])
    assert v[K_PROFIT] == 102.5
    assert v[K_RAW] == 350
    assert v[K_COMM] == -2.1 and v[K_COMM_RAW] == -12
    assert v[K_FEE] == -4 and v[K_SWAP] == -1
    assert v[K_VOL] == 3.5
    assert v[K_AGENT] == 9
    assert v[K_BAL] == 1 and v[K_CAN] == 1
    assert v[K_COL_ROWS] >= 2 and v[K_COL_LOGINS] == 2
    assert v[K_FB_ROWS] == 2 and v[K_FB_LOGINS] == 2  # #N/A rows, not a currency
    assert "30" in str(v[K_FB_LIST]) and "40" in str(v[K_FB_LIST])
    assert "10" not in str(v[K_FB_LIST]) and "20" not in str(v[K_FB_LIST])
    assert "divided by 100" in str(v[K_USC])
    assert "Client Equity FX is not" in str(v[K_USC])
print("OK header-first layout uses the currency column and divides USC by 100")

# Mapping wins over the column. Classic stays USD even when the column says USC.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src, mapping = tmp / "deals.csv", tmp / "types.csv"
    write_comma(src, [
        trade("1", login="20", profit="250", currency="USD"),
        trade("2", login="10", profit="80", currency="USC"),
    ])
    write_mapping(mapping, [("20", "CentAccount"), ("10", "Classic"),
                            ("11", "Plus"), ("12", "Pro")])
    _, hasil, _ = run([src], allow_mt4=True, mt4_accounts=mapping)
    got = by_login(hasil)
    assert got["20"]["Currency"] == "USC" and got["20"]["Profit"] == 2.5
    assert got["10"]["Currency"] == "USD" and got["10"]["Profit"] == 80
    v = verification(hasil["Deals - Daily.xlsx"]["Verifikasi"])
    assert v[K_MAP_ROWS] == 2 and v[K_MAP_LOGINS] == 2
    assert v[K_COL_ROWS] == 0
    assert "20" in str(v[K_DISAGREE]) and "10" in str(v[K_DISAGREE])
    assert v[K_FB_LIST] in ("", "(none)", None) or str(v[K_FB_LIST]) in ("(none)", "")
print("OK account-type mapping overrides the currency column")

# No column and no mapping: the per-file choice is the fallback, and those logins are listed.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "plain.csv"
    write_comma(src, [trade("1", login="77", profit="100", volume="1")], currency=False)
    _, hasil, _ = run([src], allow_mt4=True, mt4_currency={src.name: "USC"})
    got = by_login(hasil)
    assert got["77"]["Currency"] == "USC" and got["77"]["Profit"] == 1
    assert got["77"]["Volume"] == 1
    v = verification(hasil["Deals - Daily.xlsx"]["Verifikasi"])
    assert v[K_FB_ROWS] == 1 and "77" in str(v[K_FB_LIST])
    assert v[K_MAP_ROWS] == 0 and v[K_COL_ROWS] == 0
print("OK per-file USD/USC choice remains the fallback")

# MT5 USC still uses the FX rate. MT4 USC ignores that rate and divides by 100.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    mt5, mt4, fx = tmp / "mt5.csv", tmp / "mt4.csv", tmp / "fx.xlsx"
    with mt5.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Deal", "Login", "Time", "Type", "Entry", "Symbol", "Volume",
                    "Commission", "Fee", "Swap", "Profit", "Currency"])
        w.writerow(["9", "5", "2026.07.01 01:00:00", "buy", "out", "EURUSD", "1",
                    "0", "0", "0", "100", "USC"])
    write_comma(mt4, [trade("1", login="20", profit="100", currency="USC")])
    wb = Workbook()
    ws = wb.active
    ws.title = "Query result"
    ws.append(["date", "Currency", "rate"])
    ws.append([datetime.date(2026, 7, 1), "USC", 50])
    wb.save(fx)
    wb.close()
    _, _, out = run([mt5, mt4, fx], allow_mt4=True)
    wb = load_workbook(io.BytesIO(zipfile.ZipFile(out).read("Deals - Daily.xlsx")))
    try:
        headers = [c.value for c in wb["Daily"][1]]
        rows = []
        for cells in wb["Daily"].iter_rows(min_row=2, values_only=True):
            rows.append(dict(zip(headers, cells)))
    finally:
        wb.close()
    by_platform = {r["Platform"]: r for r in rows}
    assert by_platform["MT5"]["Profit"] == 100
    assert by_platform["MT5"]["Profit (USD)"] == 2
    assert by_platform["MT4"]["Profit"] == 1
    assert by_platform["MT4"]["Profit (USD)"] == 1
    assert by_platform["MT4"]["Currency"] == "USC"
print("OK MT5 USC still uses FX; MT4 USC is a flat divide by 100")

# A cp1252 comment must not stop the header-first reader. DPM still rejects the file.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "legacy.csv"
    write_comma(src, [trade("1", login="5", profit="4", comment="ok")])
    blob = src.read_bytes().replace(b"ok", b"<\xd0comment>")
    src.write_bytes(blob)
    try:
        src.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        pass
    else:
        raise AssertionError("fixture was valid utf-8")
    _, hasil, _ = run([src], allow_mt4=True)
    assert by_login(hasil)["5"]["Profit"] == 4
    try:
        proses([src], tmp / "out.zip")
    except SystemExit as exc:
        assert "MT5 Deals History only" in str(exc)
    else:
        raise AssertionError("DPM accepted a header-first MT4 file")
print("OK cp1252 header-first file; DPM still rejects it")

# Conflicting account types stop the run. An xlsx mapping is accepted.
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = tmp / "deals.csv"
    write_comma(src, [trade("1", login="20", profit="100", currency="USD")], currency=False)
    bad = tmp / "bad.csv"
    write_mapping(bad, [("20", "CentAccount"), ("20", "Classic")])
    try:
        proses([src], tmp / "out.zip", allow_mt4=True, mt4_accounts=bad)
    except SystemExit as exc:
        assert "20" in str(exc) and "account" in str(exc).lower()
    else:
        raise AssertionError("conflicting account types were accepted")
    xlsx = tmp / "types.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["account", "account_type"])
    ws.append([20, "CentAccount"])
    wb.save(xlsx)
    wb.close()
    _, hasil, _ = run([src], allow_mt4=True, mt4_accounts=xlsx)
    assert by_login(hasil)["20"]["Currency"] == "USC"
    assert by_login(hasil)["20"]["Profit"] == 1
print("OK xlsx mapping and conflicting types")

# Trimmed July export plus the trimmed mapping. Every sample login is mapped.
assert JULY.is_file() and MAPPING.is_file()
try:
    JULY.read_text(encoding="utf-8")
except UnicodeDecodeError:
    pass
else:
    raise AssertionError("July sample should not be valid utf-8")
summary, hasil, _ = run([JULY], allow_mt4=True, mt4_accounts=MAPPING)
v = verification(hasil["Deals - Daily.xlsx"]["Verifikasi"])
rows = daily_rows(hasil)
assert rows and all(r[0] == "MT4" for r in rows)
currencies = {r[11] for r in rows}
assert currencies == {"USD", "USC"}
usc_rows = [r for r in rows if r[11] == "USC"]
assert {r[1] for r in usc_rows} == {"20028179", "20033386"}
assert v[K_MAP_LOGINS] >= 2
assert v[K_FB_ROWS] == 0
assert str(v[K_FB_LIST]) in ("(none)", "")
assert v[K_DISAGREE] in ("(none)", "")
assert v[K_BAL] >= 1 and v[K_CAN] >= 1
# Sheet profit is raw USD plus USC/100. Raw profit is larger by the unconverted cents.
assert v[K_RAW] != v[K_PROFIT]
assert abs(v[K_PROFIT] - (v[K_RAW] - (v[K_RAW] - v[K_PROFIT]))) < 1e-9
print("OK trimmed July fixture resolves every login from the mapping")

# MT5-only output is unchanged by the new argument when no MT4 file is present.
_, hanya_mt5, _ = run([MT5], allow_mt4=True, mt4_accounts=MAPPING)
daily = hanya_mt5["Deals - Daily.xlsx"]["Daily"]
assert daily[0][0] == "Login"
assert "USC" in {r[10] for r in daily[1:]}
assert "MT4 USC" not in "".join(str(r[0]) for r in hanya_mt5["Deals - Daily.xlsx"]["Verifikasi"])
print("OK MT5-only workbook unchanged")


import app as A  # noqa: E402

A.app.config["TESTING"] = True
client = A.app.test_client()
client.environ_base["HTTP_AUTHORIZATION"] = "Basic " + __import__("base64").b64encode(
    b"mt4-cur:mt4-cur-password-1").decode()

kvb = client.get("/kvb/tool/segregate")
assert kvb.status_code == 200
assert b"mt4-accounts" in kvb.data
assert b"account_type" in kvb.data
assert b"CentAccount" in kvb.data
assert b"MT4 Raw Report is optional" in kvb.data
dpm = client.get("/tool/segregate")
assert b"mt4-accounts" not in dpm.data
assert b"Drop Deals History files here" in dpm.data
print("OK KVB page offers the account-type upload; DPM page does not")

body = (
    "Deal,Login,Open Time,Type,Symbol,Volume,Open Price,S/L,T/P,Close Time,Close Price,"
    "Reason,Gateway Order,Gateway Volume,Open Price Delta,Close Price Delta,Agent,"
    "Commission,Taxes,Swap,Profit,Points,Comment\n"
    "1,20,2026.07.01 00:10:00,buy,xauusd,1,,,,2026.07.01 01:00:00,,,,,,,0,0,0,0,100,,\n"
).encode()
mapping_body = b"account,account_type\n20,CentAccount\n"
commands = []
real_run = A.subprocess.run


def capture(command, **kwargs):
    commands.append(list(command))
    return real_run(command, **kwargs)


A.subprocess.run = capture
try:
    posted = client.post(
        "/kvb/tool/segregate",
        data={
            "file": (io.BytesIO(body), "deals.csv"),
            "mt4_accounts": (io.BytesIO(mapping_body), "types.csv"),
        },
        content_type="multipart/form-data")
    assert posted.status_code == 303, posted.status_code
    job_id = posted.headers["Location"].rstrip("/").rsplit("/", 1)[-1]
    state = None
    for _ in range(200):
        state = A._read_state(job_id)
        if state and state["state"] in ("done", "failed"):
            break
        __import__("time").sleep(0.02)
    assert state and state["state"] == "done", state
    assert commands and "--allow-mt4" in commands[-1]
    assert "--mt4-accounts" in commands[-1]
finally:
    A.subprocess.run = real_run
print("OK KVB job receives the account-type file")

print("OK MT4 account currency")
