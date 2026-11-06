"""MT4 account-type file with a GROUP column (account, GROUP, account_type).

USC = account_type CentAccount OR a GROUP name containing cent/cents/USC as a
word. The old account,account_type layout behaves exactly as before. All
logins and groups below are synthetic.

python3 test_deal_segregator_mt4_cent_group.py
"""
import csv
import io
import os
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(ROOT / "users.db")
os.environ["DW_JOB_DIR"] = str(ROOT / "jobs")

from deal_segregator import (_adalah_peta_akun, _group_cent, baca_peta_akun,  # noqa: E402
                             baca_peta_akun_detail, proses)

MT4_HEADER = [
    "Deal", "Login", "Open Time", "Type", "Symbol", "Volume", "Open Price",
    "S/L", "T/P", "Close Time", "Close Price", "Reason", "Gateway Order",
    "Gateway Volume", "Open Price Delta", "Close Price Delta", "Agent",
    "Commission", "Taxes", "Swap", "Profit", "Points", "Comment",
]
K_FILE = " (MT4 account types)"
K_RULE = "MT4 cent rule"
K_COUNTS = "MT4 cent counts (account_type vs GROUP)"
K_CHANGED = "MT4 logins changed to USC by GROUP (account_type alone would be USD)"
K_TYPE_ONLY = "MT4 CentAccount logins whose GROUP is not a cent group (kept USC)"

# --- group pattern ---------------------------------------------------------
for name in ("real\\Plus-10-cents", "Plus-10-cents", "Plus-10-centsFM", "real\\Pro-06-SL5-cents",
             "Test\\Test-10-cents", "Plus-3020-cents", "CENT-STD", "usc-std", "Plus cent"):
    assert _group_cent(name), name
for name in ("real\\Plus-10", "Plus-10_FM", "Percent", "KBP-KK-CME", "real\\Pro-06", "",
             None, "Plus-GAUCNY", "Accent-1", "Docents"):
    assert not _group_cent(name), name
print("OK group pattern")


def verifikasi(out):
    """Verifikasi sheet of the first workbook in the result zip, as label -> value."""
    with zipfile.ZipFile(out) as z:
        for name in sorted(z.namelist()):
            wb = load_workbook(io.BytesIO(z.read(name)), data_only=True)
            if "Verifikasi" in wb.sheetnames:
                rows = {r[0]: r[1] for r in wb["Verifikasi"].iter_rows(values_only=True) if r and r[0]}
                wb.close()
                return rows
            wb.close()
    raise AssertionError("no Verifikasi sheet")


def xlsx(path, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    return path


def csv_file(path, header, rows):
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    return path


ROWS = [
    (700001, "real\\Plus-10", "Classic"),           # USD
    (700002, "real\\Plus-10-cents", "CentAccount"),  # USC both
    (700003, "real\\Pro-06-cents", "Pro"),           # USC by group only
    (700004, "Plus-06-cents", "Pro"),                # USC by group only
    (700005, "real\\Plus-03", "Plus"),               # USD
    (700006, "real\\Plus-20", "CentAccount"),        # USC by type only (conflict shown)
    (700007, "Plus-10-centsFM", "CentAccount"),      # USC both
]

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    # New layout, the column order of the real export, xlsx and csv.
    for path in (xlsx(tmp / "MT4_Cent_Account.xlsx", ["account", "GROUP", "account_type"], ROWS),
                 csv_file(tmp / "MT4_Cent_Account.csv", ["account", "GROUP", "account_type"], ROWS)):
        assert _adalah_peta_akun(path), path
        hasil, info = baca_peta_akun_detail(path)
        assert hasil == {"700001": "USD", "700002": "USC", "700003": "USC", "700004": "USC",
                         "700005": "USD", "700006": "USC", "700007": "USC"}, hasil
        assert info["group_column"] and info["rows"] == 7
        assert (info["both"], info["group_only"], info["type_only"]) == (2, 2, 1), info
        assert info["group_only_logins"] == ["700003", "700004"]
        assert info["type_only_logins"] == ["700006"]
        assert baca_peta_akun(path) == hasil
    print("OK new layout account, GROUP, account_type (xlsx and csv)")

    # Header case and column order do not matter; Group in any case.
    p = csv_file(tmp / "order.csv", ["account_type", "Account", "Group"],
                 [("Pro", "700010", "real\\Pro-06-SL5-cents"), ("Classic", "700011", "Plus-10")])
    assert baca_peta_akun(p) == {"700010": "USC", "700011": "USD"}
    print("OK any column order and header case")

    # Old layout: unchanged behaviour, no group statistics.
    old_rows = [(700001, "Classic"), (700002, "CentAccount"), (700003, "Pro"), (700004, "cent account")]
    for path in (xlsx(tmp / "old.xlsx", ["account", "account_type"], old_rows),
                 csv_file(tmp / "old.csv", ["account", "account_type"], old_rows)):
        hasil, info = baca_peta_akun_detail(path)
        assert hasil == {"700001": "USD", "700002": "USC", "700003": "USD", "700004": "USC"}
        assert not info["group_column"] and info["group_only"] == 0 and info["type_only"] == 0
    print("OK old layout account, account_type unchanged")

    # Same login twice: same currency is fine, a different one stops the run.
    p = csv_file(tmp / "dup.csv", ["account", "GROUP", "account_type"],
                 [(700020, "Plus-10-cents", "Pro"), (700020, "Plus-20-cents", "CentAccount")])
    hasil, info = baca_peta_akun_detail(p)
    assert hasil == {"700020": "USC"} and info["repeated"] == 1
    p = csv_file(tmp / "dup2.csv", ["account", "GROUP", "account_type"],
                 [(700021, "Plus-10-cents", "Pro"), (700021, "Plus-10", "Pro")])
    try:
        baca_peta_akun(p)
    except SystemExit as e:
        assert "conflicting account types" in str(e)
    else:
        raise AssertionError("conflicting duplicate must stop")
    print("OK duplicate logins")

    # End to end: group-only cent login is divided by 100, Verifikasi explains why.
    mt4 = tmp / "MT4_RawReport_20260712.csv"
    with mt4.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(MT4_HEADER)
        for deal, login in (("1", "700001"), ("2", "700003"), ("3", "700002")):
            row = {"Deal": deal, "Login": login, "Open Time": "2026.07.12 00:10:00",
                   "Type": "buy", "Symbol": "xauusd", "Volume": "1",
                   "Close Time": "2026.07.12 01:00:00", "Agent": "0", "Commission": "-2",
                   "Taxes": "0", "Swap": "0", "Profit": "100", "Comment": ""}
            w.writerow([row.get(h, "") for h in MT4_HEADER])
    accounts = xlsx(tmp / "MT4_Cent_Account_e2e.xlsx", ["account", "GROUP", "account_type"], ROWS)
    out = tmp / "out.zip"
    proses([mt4], out, allow_mt4=True, mt4_accounts=accounts)
    summary = verifikasi(out)
    assert summary["MT4 Total Profit (closed buy/sell)"] == 102.0, summary["MT4 Total Profit (closed buy/sell)"]
    line = summary["  - " + accounts.name + K_FILE]
    assert "columns account, GROUP, account_type" in line and "5 USC, 2 USD" in line, line
    assert "CentAccount" in summary[K_RULE] and "Plus-10-cents" in summary[K_RULE]
    counts = summary[K_COUNTS]
    assert "both CentAccount and cent group: 2" in counts
    assert "cent group only (type not CentAccount, USC by group): 2" in counts
    assert "CentAccount only (group not cent, USC by type): 1" in counts
    assert "neither (USD): 2" in counts
    changed = summary[K_CHANGED]
    assert "2 logins in the file, 1 of them in this run's MT4 rows" in changed, changed
    assert "Logins in this run: 700003." in changed and "All logins: 700003, 700004" in changed
    assert "real\\Pro-06-cents / Pro: 1" in changed
    assert "700006" in summary[K_TYPE_ONLY]
    assert "cent GROUP" in summary["  - " + mt4.name + " (MT4 Raw Report)"]
    assert "GROUP column" in summary["MT4 currency rule"]

    # Same run with the old layout: 700003 is USD, no group rows on Verifikasi.
    old = csv_file(tmp / "types_old.csv", ["account", "account_type"],
                   [(a, t) for a, _g, t in ROWS])
    proses([mt4], tmp / "old.zip", allow_mt4=True, mt4_accounts=old)
    summary_old = verifikasi(tmp / "old.zip")
    assert summary_old["MT4 Total Profit (closed buy/sell)"] == 201.0
    assert "3 CentAccount (USC), 4 other types (USD)" in summary_old["  - " + old.name + K_FILE]
    assert K_RULE not in summary_old and K_CHANGED not in summary_old
    print("OK end to end: GROUP cent is ÷100 and listed on Verifikasi; old layout unchanged")

    # Account file uploaded among the regular files (the browser now sends it in chunks).
    proses([mt4, accounts], tmp / "mixed.zip", allow_mt4=True)
    summary_mixed = verifikasi(tmp / "mixed.zip")
    assert summary_mixed["MT4 Total Profit (closed buy/sell)"] == 102.0
    print("OK account file found among the uploaded files")

print("OK mt4 cent group")
