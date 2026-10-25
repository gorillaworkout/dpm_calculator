#!/usr/bin/env python3
"""dw-calculator web UI. Landing page + menu generator.

Jalankan:  python3 app.py     -> http://127.0.0.1:5000
Menu baru: tambahkan entry di TOOLS. Tidak perlu ubah kode lain.
"""
import datetime
import errno
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
from collections import defaultdict, deque
from pathlib import Path

from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   send_file, url_for)
from markupsafe import escape
from werkzeug.exceptions import HTTPException
from werkzeug.utils import secure_filename

import angka
import auth
import chunked
import cleanup

BASE = Path(__file__).resolve().parent
JOBS = Path(os.environ.get("DW_JOB_DIR", Path(tempfile.gettempdir()) /
                           "dw-calculator-jobs")).expanduser().resolve()
JOBS.mkdir(exist_ok=True)


# --------------------------------------------------------------------- version
# The server has no .git (deploys copy single files from raw GitHub at a SHA),
# so the running version comes from the VERSION file, read once at startup.
# Bump VERSION in every PR; see the header of that file.
APP_NAME = "DW Calculator"


def baca_versi(path=None):
    """{'version', 'build'} from VERSION. Missing or unreadable -> 'unknown'.

    Accepts `key=value` lines (`#` comments allowed) or a bare version on the
    first non-comment line.
    """
    path = Path(path) if path else BASE / "VERSION"
    hasil = {"version": "unknown", "build": ""}
    try:
        teks = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return hasil
    for baris in teks.splitlines():
        baris = baris.strip()
        if not baris or baris.startswith("#"):
            continue
        if "=" in baris:
            kunci, _, nilai = baris.partition("=")
            kunci = kunci.strip().lower()
            if kunci in hasil and nilai.strip():
                hasil[kunci] = nilai.strip()
        elif hasil["version"] == "unknown":
            hasil["version"] = baris
    return hasil


def label_versi(info):
    """'DW Calculator v1.28 · 2026-10-25' (no date part when build is empty)."""
    v = info.get("version") or "unknown"
    teks = f"{APP_NAME} v{v}" if v != "unknown" else f"{APP_NAME} version unknown"
    return f"{teks} · {info['build']}" if info.get("build") else teks


VERSI = baca_versi()
VERSI_LABEL = label_versi(VERSI)


# Job state lives in a small JSON file NEXT TO the job directory, not inside it.
# gunicorn runs several workers and each poll can land on a different one, so an
# in-memory dict returns 404 at random. Keeping the marker outside the directory
# also means the heavy result can be deleted the moment it is downloaded while
# the status page still knows what happened -- a bare 404 reads as a crash.
def _state_path(job_id):
    return JOBS / f"{job_id}.json"


def _job_dir(job_id):
    return JOBS / job_id


def _write_state(job_id, **fields):
    p = _state_path(job_id)
    p.parent.mkdir(parents=True, exist_ok=True)
    state = _read_state(job_id) or {}
    state.update(fields)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(state))
    tmp.replace(p)                      # atomic: a poll never sees a half-written file
    return state


def _read_state(job_id):
    try:
        return json.loads(_state_path(job_id).read_text())
    except (OSError, ValueError):
        return None

# slug -> (judul, deskripsi, langkah, siap?)
#   langkah = tuple nama script, dijalankan BERURUTAN. Output langkah pertama jadi
#             input langkah kedua. Tiap script dipanggil: script <input> -o <output>.
# Sejak 24 Aug 2026 MTOATD ikut dihitung di hitung_dw.py -- SEKALI JALAN, tidak ada
# menu tahap kedua lagi. hitung_mtoatd.py sudah dihapus.
# Both companies export MT5 Deals History with the same default template, so Deal
# Segregator is one shared definition. Its script list is not the D&W pipeline.
_SEGREGATE = ("Deal Segregator", "Combine MT5 Deals History files and optionally add "
              "a Client Equity FX workbook for USD-converted financial totals.",
              ("deal_segregator.py",), True)
TOOLS = {
    # SATU menu saja (permintaan user 26 Aug 2026) supaya tidak ada yang bingung.
    # isi_template.py memindahkan D, W, Fund Transfer Table, Opening Balance, kurs
    # dan tabel fee dari file yang diupload ke template bersih, lalu hitung_dw.py
    # menghitung semuanya. Tidak ada rate/kurs yang disimpan di tool ini -- semua
    # dibaca dari file yang diupload, karena rate bisa berubah kapan saja.
    "dw": ("Generate D&W", "Upload your D&W workbook - the raw back-office export or a "
           "Template D&W you filled in yourself. Everything is produced in a single run: "
           "D&W Report, D&W Detail, D&W FEE, Channel Balance, J Wallet, MTOATD, "
           "Missing Data, Legend. Fee rates and Xero rates are always read from the file "
           "you upload, never stored here.",
           ("isi_template.py", "hitung_dw.py"), True),
    "segregate": _SEGREGATE,
    # Own branch in run(): two uploads, no month dropdown, .xlsx output.
    "pl-desk": ("PL by Desk", "Monthly P&L by sales desk. Upload the FinanceOS "
                "settlement PDF and the Metabase desk-rebate workbook.",
                ("pl_desk.py",), True),
    # DPM only. A month of daily segregated files + Client Equity FX -> one .xlsx.
    "segregated": ("Segregated", "Monthly roll-up of the daily segregated files: upload "
                   "the whole month plus the month's Client Equity FX file. One workbook "
                   "with the monthly sum per account (Balance and Equity from the last day), "
                   "every daily row, and USD columns.",
                   ("dpm_segregated.py",), True),
}

# Separate registry: KVB never changes or extends the existing DPM pipeline.
KVB_TOOLS = {
    "dw": ("Generate D&W", "Upload a KVB Plus workbook and generate the finished monthly "
           "D&W report in one run.",
           ("siapkan_kvb.py", "isi_template.py", "hitung_dw.py"), True),
    "segregate": ("Deal Segregator",
                  "Combine MT5 Deals History files. MT4 Raw Report is optional. "
                  "You can also add a Client Equity FX workbook for USD-converted financial totals.",
                  ("deal_segregator.py",), True),
}


def _tools_for(company):
    return TOOLS if company == "dpm" else KVB_TOOLS if company == "kvb" else None


def _browser_page():
    """True when a browser navigation should not see a bare 403 page.

    Address-bar and link loads send Accept with text/html. Accept */* (curl
    and fetch) and a missing Accept stay on a plain 401 or 403, so a script
    is not handed the login HTML. The chunked uploader only POSTs, so this
    choice does not affect it. Job downloads stay 403 even for a browser.
    """
    if request.method != "GET" or request.path.endswith("/download"):
        return False
    return "text/html" in (request.headers.get("Accept") or "")


def _page_on_company(path, company):
    """The same page on the other company, or None when that page does not exist.

    Job URLs are never rewritten. Sending someone to the other company's job
    address would be a way to reach another account's result.
    """
    if path == "/kvb":
        rest = "/"
    elif path.startswith("/kvb/"):
        rest = path[4:]
    else:
        rest = path
    if rest == "/":
        return "/kvb" if company == "kvb" else "/"
    if rest.startswith("/job/"):
        return None
    parts = rest.strip("/").split("/")
    if len(parts) == 2 and parts[0] == "tool":
        slug = parts[1]
        registry = _tools_for(company)
        if not registry or slug not in registry or not registry[slug][3]:
            return None
        prefix = "/kvb" if company == "kvb" else ""
        return f"{prefix}/tool/{slug}"
    if rest == "/template":
        return "/kvb/template" if company == "kvb" else "/template"
    return None


def _redirect_denied(company):
    """Send a browser somewhere it is allowed to be, and say why."""
    other = "kvb" if company == "dpm" else "dpm"
    user = getattr(request, "user", None) or auth.current_user()
    if auth.has_company(user, other):
        dest = _page_on_company(request.path, other)
        if dest and dest != request.path:
            flash(f"Your account does not have access to {company.upper()}. "
                  f"Opened the {other.upper()} page instead.")
            return redirect(dest)
        home = "/kvb" if other == "kvb" else "/"
        if home != request.path:
            flash(f"Your account does not have access to {company.upper()}.")
            return redirect(home)
    flash(f"Your account does not have access to {company.upper()}. "
          f"Please sign in with a {company.upper()} account.")
    return redirect(url_for("login", next=request.full_path.rstrip("?")))


def _require_company(company, browser_redirect=True):
    if auth.has_company(request.user, company):
        return None
    if browser_redirect and _browser_page():
        return _redirect_denied(company)
    abort(403, description=f"Your account does not have access to {company.upper()}.")


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024

# Every route is behind HTTP Basic Auth. A before_request hook is used instead of
# decorating each view so a new endpoint cannot be added unprotected by accident.
auth.register(app)
chunked.register(app, JOBS)


@app.teardown_request
def _lepas_kunci_job(_exc=None):
    # A run that returned an error before its thread started still holds its
    # job lock. Release it so the next cleanup can remove the leftovers.
    fh = g.pop("job_lock", None)
    if fh is not None:
        cleanup.release_job_lock(fh)


@app.before_request
def _require_login():
    if request.endpoint in ("static", "login", "login_post", "logout"):
        return None
    if request.path.startswith("/admin"):
        return None                      # admin views carry their own stricter check
    user = auth.current_user()
    if not user:
        # Browsers get the styled sign-in page. Accept */* and a missing Accept
        # stay a plain 401: that is curl, fetch, and other scripts. The chunked
        # uploader only POSTs, so it never follows this redirect.
        if _browser_page():
            return redirect(url_for("login", next=request.full_path.rstrip("?")))
        return auth._deny()
    request.user = user
    if (request.path == "/" or request.path.startswith("/tool/") or
            request.path == "/template"):
        return _require_company("dpm")
    if request.path == "/kvb" or request.path.startswith("/kvb/"):
        return _require_company("kvb")
    return None


# ----------------------------------------------------------------- report month
# Every upload must say WHICH MONTH it is a report for. A back-office export for
# June routinely runs from late May to early July: the export is filtered on
# Apply Date, while the report itself is built on Paid Date (deposits) and
# Completed Date (withdrawals). Without a month, those tails end up in the
# totals, in the channel balances and in the MTOATD date blocks.
#
# The chosen month is handed to each step under its own flag name; a step that
# is not listed simply does not get it.
PERIODE_ARG = {"isi_template.py": "--bulan", "hitung_dw.py": "--period"}

# Saldo penutup J Wallet bulan SEBELUMNYA, diketik di halaman upload. Hanya
# hitung_dw.py yang memakainya. Tanpa ini J Wallet mulai dari NOL setiap kali the
# uploaded workbook has no 'J Wallet' / 'Opening Balance' sheet - which is the
# normal case, so the closing balance never carried over from one month to the next.
SALDO_ARG = {"hitung_dw.py": "--jwallet-opening"}

# Saldo pembuka per Payment Channel + Currency. Isinya BANYAK BARIS (40-70), jadi
# tidak bisa lewat satu field seperti J Wallet: orang menempel tabelnya, kita simpan
# jadi file di folder job, lalu path-nya diberikan ke hitung_dw.py.
CHANNEL_ARG = {"hitung_dw.py": "--channel-opening"}
# KVB tidak boleh memakai EXTRA_FEES (rate DPM yang ditulis di kode). Flag harus
# ikut sejak isi_template.py karena tahap itu bisa menggabungkan beberapa fee sheet.
# Setiap nilai adalah tuple argumen CLI, hanya untuk company KVB.
# DPM tidak menerima flag ini.
KVB_ARG = {
    "isi_template.py": ("--tanpa-extra-fees",),
    "hitung_dw.py": (
        "--tanpa-extra-fees",
        "--usd-shadow", "CNY:CHIPPAY,EP,UPAY,ZPAY",
        "--fee-currency-suffix",
        "--warn-high-pct",
        "--cny-rate-check",
    ),
}
# The typed balance is read by angka.parse_number in THOUSANDS mode, exactly as
# hitung_dw.baca_angka_saldo reads it ('3.5E+09', 'Rp 3.500.000,50', '(1,234)').
SALDO_MAKS_PANJANG = 40
PERIODE_POLA = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
NAMA_BULAN = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]


def _bulan_sebelum(bulan, tahun):
    """(6, 2026) -> ('May', 2026). Dipakai untuk memberi label field saldo pembuka."""
    b, t = (12, tahun - 1) if bulan == 1 else (bulan - 1, tahun)
    return NAMA_BULAN[b - 1], t


def _pilihan_bulan():
    """Isi dropdown bulan + tahun, plus bulan lalu sebagai default.

    The default is the PREVIOUS month: a monthly report is always run after the
    month has closed, so that is the one being asked for nearly every time.
    """
    hari_ini = datetime.date.today()
    lalu = (hari_ini.replace(day=1) - datetime.timedelta(days=1))
    sb, st = _bulan_sebelum(lalu.month, lalu.year)
    return {
        "bulan_opsi": list(enumerate(NAMA_BULAN, start=1)),
        "tahun_opsi": list(range(hari_ini.year - 3, hari_ini.year + 2)),
        "bulan_default": lalu.month,
        "tahun_default": lalu.year,
        "nama_bulan": NAMA_BULAN,
        "saldo_label": f"{sb} {st}",
        "saldo_default": "",
        "saldo_ch_default": "",
    }


def _baca_periode(form):
    """-> ('2026-06', None) atau (None, 'pesan error')."""
    bulan, tahun = (form.get("bulan") or "").strip(), (form.get("tahun") or "").strip()
    if not bulan or not tahun:
        return None, "Please choose the report month before uploading."
    nilai = f"{tahun}-{int(bulan):02d}" if bulan.isdigit() and tahun.isdigit() else ""
    if not PERIODE_POLA.match(nilai):
        return None, f"That report month is not valid: {tahun}-{bulan}."
    return nilai, None


def _label_periode(periode):
    """'2026-06' -> 'Jun 2026' (dipakai di nama file hasil)."""
    th, bl = periode.split("-")
    return f"{NAMA_BULAN[int(bl) - 1][:3]} {th}"


# Short hint shown next to the file picker, per menu.
HINT = {
    "dw": "Must have the transaction sheets (D / W, or Deposits / Withdrawals), a fee "
          "table ('D&W FEE', or 'D&W TD FEE' + '-add.') and a 'Xero' sheet with the "
          "rates. Sheets 'Fund Transfer Table', 'Opening Balance', 'J Wallet' and "
          "'Payment Channel Balance' are used when present, so that Channel Balance and "
          "J Wallet come out complete.",
    "segregate": "Choose one or more MT5 Deals History files (.csv/.xlsx/.xlsm). You may "
                 "also include one Client Equity FX / rate file (.xlsx or .csv, any sheet name, "
                 "columns date/currency/rate, headers not case-sensitive); "
                 "the two file types are detected automatically.",
    "pl-desk": "Upload the FinanceOS settlement (PDF, or the converted Excel) and the "
               "Metabase workbook. The workbook must contain the desk-rebate export, "
               "the Desks sheet and the Target sheet. File types are detected by content.",
    "segregated": "Select every daily segregated file of the month (.csv/.xlsx/.xlsm, one per "
                  "day) and the month's Client Equity FX file (columns date/currency/rate). "
                  "The FX file is required.",
}


# The 14 sheets a full calculation produces. Reused by the pages that produce all of them.
SHEET_PENUH = [
    ("D&W Report", "Summary by Currency x Payment Gateway — the deposit and withdrawal "
                   "totals, FX gain/loss, and CRM vs Xero average rates."),
    ("Guide", "Short in-file explanation of what goes where."),
    ("D", "Deposit transactions, with the four calculated columns added on the right."),
    ("W", "Withdrawal transactions, same four calculated columns."),
    ("Xero", "The daily exchange rates used for the calculation."),
    ("Fund Transfer Table", "Money movements between departments and channels."),
    ("Opening Balance", "Starting balances per channel and for the J Wallet."),
    ("Channel Balance", "Date x Currency x Channel with a running balance."),
    ("J Wallet (calc)", "USDT wallet ledger, derived from the Fund Transfer Table."),
    ("MTOATD", "Daily MT4 / Wallet / CRM / TD reconciliation. When the "
               "'MT4+Wallet-CRM' row reads 0 across every currency, it balances."),
    ("D&W FEE", "The merged fee table, plus a note on which rate overrode which."),
    ("D&W Detail", "D and W combined, one row per transaction."),
    ("Missing Data", "Anything missing or needing confirmation. Read this first."),
    ("Legend", "Explanation of the columns and the cell colours."),
]

CATATAN_MISSING = (
    "<strong>Always open the “Missing Data” sheet first.</strong> If a Xero rate or a fee "
    "rate could not be found, those cells are left red and nothing is calculated for them. "
    "The report will still look complete, but the totals will be short. When every line "
    "there reads “None”, the run is clean."
)

CATATAN_ASLI = ("Your uploaded file is never modified. The result is always a new file, "
                "downloaded straight to your browser.")


# Full per-page documentation, rendered under the upload form.
DOC = {
    "dw": {
        "judul": "Generate D&W — the one-step monthly run",
        "ringkas": "This is the only page you need. You give it your D&W workbook and it "
                   "returns the finished report. Internally it runs two steps back to "
                   "back — prepare, then calculate — so there is nothing else to visit.",
        "siapkan": [
            "Your D&amp;W workbook as <code>.xlsx</code> or <code>.xlsm</code>. Either "
            "form is accepted: the <strong>raw export from back office</strong>, or a "
            "<strong>Template D&amp;W you filled in yourself</strong>. Do not clean up a "
            "raw export first — it is read as-is.",
            "The transaction sheets must be present. They may be named <code>D</code> and "
            "<code>W</code>, or <code>Deposits</code> and <code>Withdrawals</code> — both "
            "are recognised.",
            "A fee table must be in the same workbook — either <code>D&amp;W FEE</code>, "
            "or <code>D&amp;W TD FEE</code> plus <code>D&amp;W TD FEE -add.</code> when "
            "it is split in two. Where the two disagree, the additional table wins and "
            "the override is recorded in the result.",
            "A <code>Xero</code> sheet holding the daily rates for the period. Rates for "
            "dates that are not in it cannot be resolved, and those rows stay red.",
            "Optional but recommended: sheets <code>Fund Transfer Table</code>, "
            "<code>Opening Balance</code>, <code>J Wallet</code> and <code>Payment "
            "Channel Balance</code>. When these are present, Channel Balance and J Wallet "
            "come out fully populated. Without them those balances simply start from zero.",
            "The <strong>J Wallet opening balance</strong> &mdash; the closing balance of "
            "the month before. Optional, but without it the J Wallet ledger restarts at "
            "zero every month, because the balance is not derivable from the "
            "transactions. If the workbook you upload already carries an "
            "<code>Opening Balance</code> or <code>J Wallet</code> sheet, leave the field "
            "empty and those are used instead. Channel balances are separate and still "
            "need those sheets.",
            "The <strong>report month</strong>, chosen on this page. It is required. A "
            "back-office export for June routinely runs from late May to early July "
            "&mdash; the export is filtered on Apply Date while this report is built on "
            "Paid Date and Completed Date &mdash; and those extra days would otherwise "
            "land in the totals, the channel balances and the MTOATD date blocks.",
            "Nothing needs to be typed in by hand. Every calculated column is written "
            "by the tool, and anything already sitting in those columns is overwritten.",
        ],
        "langkah": [
            "<strong>Step 0 — the month.</strong> Every row is judged on the date the "
            "report itself uses: Paid Date for deposits, Completed Date for "
            "withdrawals. Rows dated outside the month you chose are dropped, in "
            "<code>D</code>, in <code>W</code> and in the <code>Fund Transfer "
            "Table</code>. They stay visible in D and W with their four calculated "
            "columns blank and shaded grey, and the count is listed in section 0 of "
            "<code>Missing Data</code>. If nothing at all falls inside that month, the "
            "run stops and tells you which months the file does hold.",
            "<strong>Step 1 — prepare.</strong> Sheets D, W, Fund Transfer Table and "
            "Opening Balance are copied out of your workbook into a clean Template "
            "D&amp;W, along with the fee table and Xero rates. Opening balances are read "
            "from <code>J Wallet</code> and <code>Payment Channel Balance</code> when "
            "those sheets exist.",
            "<strong>Step 2 — calculate.</strong> Handling Fee, Xero Rate, Xero USD and "
            "Forex Gain/Loss are computed row by row, then every report sheet is built: "
            "D&amp;W Report, D&amp;W Detail, D&amp;W FEE, Channel Balance, J Wallet "
            "(calc), MTOATD, Missing Data and Legend.",
            "The finished workbook downloads automatically. A typical month takes only a "
            "few seconds.",
        ],
        "hasil": "One <code>.xlsx</code> file with 14 sheets:",
        "sheets": SHEET_PENUH,
        "catatan": [CATATAN_MISSING, CATATAN_ASLI],
        "catatan_penting": True,
    },
    "segregate": {
        "judul": "Deal Segregator — daily and monthly trading summaries",
        "ringkas": "Upload one or more MT5 Deals History exports, optionally together with a "
                   "Client Equity FX workbook. File types are detected automatically.",
        "siapkan": [
            "One or more MT5 Deals History files in <code>.csv</code>, <code>.xlsx</code> or <code>.xlsm</code> format.",
            "Each file (or each sheet of an Excel file) must contain Deal, Login, Time, Type, Entry, Symbol, Volume, Commission, Fee, Swap, Profit and Currency columns. Group and Country are ignored and are not required.",
            "Optional: one <strong>Client Equity FX</strong> / rate file in <code>.xlsx</code> "
            "or <code>.csv</code> format with columns <code>date</code>, <code>currency</code>, "
            "<code>rate</code>. Any sheet name is fine and headers are not case-sensitive. "
            "Select it together with the Deals files.",
            f"There is a <strong>{chunked.teks_batas()} total upload limit</strong> and a "
            f"<strong>{chunked.MAX_FILES}-file maximum</strong> per run. Large batches are sent to the server "
            "in pieces automatically &mdash; you will see "
            "the progress under the button, so leave the tab open until it says "
            "<em>Processing</em>. Duplicate Deal IDs are removed within each run.",
        ],
        "langkah": [
            "Rows from every file, and from every sheet of an Excel file, are pooled. "
            "<code>Entry = out</code> rows (closing trades) and <code>Entry = in</code> rows (opening trades) "
            "are kept on <strong>separate sheets</strong> with the same layout. Nothing is merged between "
            "them: each row keeps its own Volume, Commission, Fee, Swap and Profit exactly as MT5 reports it. "
            "The <code>in</code> and <code>out</code> of one trade share the same Volume, so do not add the "
            "Volume of the two sheets together.",
            "Repeated non-empty Deal IDs are removed, then Daily and Monthly Summary totals are grouped by account, period, type, symbol and currency. Group and Country are deliberately ignored.",
            "When Client Equity FX is included, Commission, Fee, Swap and Profit are converted "
            "to USD by exact Date + Currency. Missing rates stay blank and are highlighted yellow.",
        ],
        "hasil": "One <code>.zip</code> containing two workbooks:",
        "sheets": [("Deals - Daily.xlsx", "<code>Daily</code> (out rows), <code>Daily - In</code> (in rows), USD columns when FX is supplied, plus Verifikasi."),
                   ("Deals - Monthly Summary.xlsx", "<code>Monthly Summary</code> (out rows), <code>Monthly Summary - In</code> (in rows), USD columns when FX is supplied, plus Verifikasi.")],
        "catatan": ["Review the <strong>Verifikasi</strong> sheet before using the totals.",
                    "After a successful run the uploaded file is deleted from the server. "
                    "The file on your computer is not changed. Verifikasi records this.",
                    CATATAN_ASLI],
    },
    "segregated": {
        "judul": "Segregated — the month's daily files in one workbook",
        "ringkas": "Upload all daily segregated files of one month at once, together with "
                   "that month's Client Equity FX file. You get one <code>.xlsx</code>: the "
                   "monthly total per account and every daily row underneath it.",
        "siapkan": [
            "The <strong>daily segregated files</strong> of the month, one per day, in "
            "<code>.csv</code>, <code>.xlsx</code> or <code>.xlsm</code>. Select the whole "
            "month (30 or 31 files) in one go.",
            "Columns are found by their header name (not case-sensitive), so the column "
            "order does not matter and new columns are picked up automatically. Required: an "
            "account column (<code>Login</code> or <code>Account</code>). Expected: "
            "<code>Currency</code>, <code>Balance</code>, <code>Equity</code> and the amount "
            "columns.",
            "The day of each file comes from its <code>Date</code> column; without one, from "
            "the file name (for example <code>2026-09-01</code> or <code>20260901</code>).",
            "<strong>Required:</strong> the month's <strong>Client Equity FX</strong> file "
            "(<code>date</code>, <code>currency</code>, <code>rate</code>), the same file "
            "Deal Segregator uses. The run stops without it.",
            f"There is a <strong>{chunked.teks_batas()} total upload limit</strong> and a "
            f"<strong>{chunked.MAX_FILES}-file maximum</strong> per run. Files are sent in "
            "pieces automatically; leave the tab open until it says <em>Processing</em>.",
        ],
        "langkah": [
            "Every numeric column is <strong>summed over the month</strong> per account, "
            "except <code>Balance</code> and <code>Equity</code>: those take the "
            "<strong>last day's</strong> value. Prices, rates and levels (for example "
            "<code>Margin Level %</code>) are not summed and appear only in Daily.",
            "USD: each day's amount is divided by that day's Client Equity FX rate for the "
            "account's currency, then summed. Balance and Equity use the last day's value "
            "at the last day's rate. A missing rate leaves the USD cell blank and yellow.",
            "Accounts missing from the last day's file, and Balance or Equity blank or 0 "
            "on the last day, are highlighted yellow and explained in the "
            "<code>Check</code> column.",
        ],
        "hasil": "One <code>.xlsx</code> file with three sheets:",
        "sheets": [("Monthly Summary", "One row per account: the monthly sums, Balance and "
                    "Equity from the last day, the USD columns and a Check column."),
                   ("Daily", "Every daily row stacked, with a Date column and the USD columns. "
                    "Continues on Daily 2, ... past Excel's row limit."),
                   ("Verifikasi", "Files read, how each column was treated, and every check: "
                    "missing FX rates, duplicate dates, missing days, accounts missing on the "
                    "last day, Balance/Equity blank or 0.")],
        "catatan": ["Review the <strong>Verifikasi</strong> sheet before using the totals.",
                    "After a successful run the uploaded files are deleted from the server. "
                    "The files on your computer are not changed.",
                    CATATAN_ASLI],
    },
    "pl-desk": {
        "judul": "PL by Desk — monthly P&L by sales desk",
        "ringkas": "Reconcile the FinanceOS settlement with the Metabase desk-rebate export "
                   "and split the Metabase totals across the sales desks.",
        "siapkan": [
            "The <strong>FinanceOS settlement</strong> for one calendar month, "
            "<code>CONSOLIDATED</code> USD, for DPM Malaysia. PDF is the primary input. "
            "The team's converted Excel is accepted as a fallback and runs through the same checks.",
            "The <strong>Metabase workbook</strong> for that same month: the "
            "<code>Query result</code> sheet (one row per country), plus <code>Desks</code>, "
            "<code>Target</code> and <code>Country code</code>. Desk mapping and targets are "
            "read from this upload on every run.",
            "The two files are recognised by their contents, so the names do not matter. "
            "A scanned PDF is refused.",
        ],
        "langkah": [
            "The settlement text layer is read and checked against its own sub-totals, "
            "equity block and required labels. A break stops the run and no workbook is produced.",
            "Metabase totals land in column C. FinanceOS amounts land in column B, except the "
            "yellow rows, which stay blank. Variance is column B minus column C.",
            "Countries missing from the Desks sheet are assigned to <strong>GEM</strong> and "
            "listed on the Check sheet. Desk Actual columns are SUMIFS formulas, and "
            "Completion % is Actual / Target.",
            "Orange manual-rebate rows are left blank for key-in. Their group total is a live SUM.",
        ],
        "hasil": "One <code>.xlsx</code> workbook:",
        "sheets": [
            ("Outcome", "The monthly report. Sub-totals, variance, NDP, gross profit, desk Actual and Completion % are Excel formulas."),
            ("FinanceOS (parsed)", "Every settlement line, with the reconciliation checks."),
            ("Raw Metabase", "The export plus <code>desk</code> and <code>country_name</code>."),
            ("Desks", "Copied from the upload."),
            ("Target", "Copied from the upload."),
            ("Country code", "Copied from the upload."),
            ("Check", "Rounding, Diff, wrapped labels, new labels, test entries, GEM assignments, variances over 1,000, blank targets, and the desks-minus-total check."),
        ],
        "catatan": [
            "Open the <strong>Check</strong> sheet first. Cash Movement column B is left blank. "
            "Deposit and Withdrawal are compared with Metabase in the block under Gross profit.",
            CATATAN_ASLI,
        ],
    },
}

KVB_HINT = ("Upload one KVB Plus workbook (.xlsx/.xlsm) containing Transfer, D, W, Xero, "
            "and Handing Fee or Handling Fee sheets.")
KVB_DOC = {
    "judul": "KVB Generate D&W — one monthly run",
    "ringkas": "Upload the KVB Plus workbook. The app translates KVB fields, prepares a clean D&W template, then calculates the finished report.",
    "siapkan": [
        "One KVB Plus workbook with <code>Transfer</code>, <code>D</code>, <code>W</code>, <code>Xero</code>, and <code>Handing Fee</code> or <code>Handling Fee</code> sheets.",
        "The report month. The uploaded workbook is never modified.",
    ],
    "langkah": [
        "KVB sheet and field names are translated without changing transaction amounts.",
        "The translated data is copied into the clean D&amp;W template.",
        "The report is calculated and returned as a new workbook.",
        "CNY gateways ChipPay, EP, UPay and ZPay also get a USD block directly under "
        "the CNY block on <code>Payment Channel Balance</code>. Deposit and Withdrawal "
        "on that block come from the USD column. Charges are Handling Fee × (USD ÷ Transaction). "
        "CNY fund transfers, outgoing and incoming, and the sending-side transfer fee, "
        "are converted onto that USD block. The rate is the Xero USD when this run "
        "computed it, otherwise that gateway's same-day weighted USD/Transaction, "
        "otherwise its monthly weighted ratio, otherwise the run's CNY median ratio. "
        "Legend names the rate used for each transfer.",
    ],
    "hasil": "One finished <code>.xlsx</code> D&amp;W workbook.",
    "sheets": SHEET_PENUH,
    "catatan": [
        CATATAN_MISSING,
        "Fee rates on the Handling Fee sheet: <code>+8</code> is a flat 8 per transaction "
        "(the plus sign is required). <code>1.5%+50</code> is 1.5% plus a flat 50. "
        "<code>0.45%,min 90</code> is 0.45% with a minimum of 90. A plain number "
        "<code>8</code> is read as 800%, not as a flat fee. Write <code>+8</code> when "
        "you mean a flat amount. The calculation is not rewritten when a rate reads as "
        "100% or more; Missing Data names those cells.",
        "USD opening balance for a CNY gateway, paste exactly one of these: "
        "<code>ChipPay USD(CNY) 12,345.67</code>, or three columns "
        "<code>ChipPay</code>, <code>USD(CNY)</code>, <code>12,345.67</code> "
        "(tab or two spaces). <code>USD (CNY)</code> with a space is the same. "
        "That is the USD block under CNY ChipPay, not a real USD ChipPay balance "
        "(that one stays <code>ChipPay</code> / <code>USD</code> / amount). "
        "Leave the line out and the USD block starts at zero.",
        "A withdrawal gateway <code>PA</code> in currency <code>PKR</code> uses the fee "
        "row <code>PA-PKR</code> when no exact <code>PA</code> row exists. An exact row "
        "still wins. Missing Data records it as <code>PA (PKR) used fee row PA-PKR</code>.",
        CATATAN_ASLI,
    ],
    "catatan_penting": True,
}


KVB_SEGREGATE_HINT = (
    "Choose one or more MT5 Deals History files (.csv/.xlsx/.xlsm). "
    "MT4 Raw Report is optional and is detected from the file contents, not the file name. "
    "Any MT4 file requires the account-type file (account, account_type): CentAccount = USC (÷100), other types = USD. "
    "You may also include one Client Equity FX workbook (.xlsx, sheet 'Query result')."
)
KVB_SEGREGATE_DOC = {
    "judul": "Deal Segregator — MT5, with optional MT4",
    "ringkas": "Upload one or more MT5 Deals History exports. MT4 Raw Report is optional. "
               "A Client Equity FX workbook is optional too. File types are detected from "
               "their contents, not from the file name.",
    "siapkan": [
        "One or more MT5 Deals History files in <code>.csv</code>, <code>.xlsx</code> or <code>.xlsm</code> format.",
        "Optional: one or more MT4 files. Either a Manager <strong>Raw Report</strong> "
        "(semicolon-separated, starting with <code>Raw Report for ...</code>) or a header-first export "
        "(comma-separated, with or without a trailing currency column). "
        "MT4 is optional. If you do not upload one, the result is the same MT5 workbook as before.",
        "<strong>Required with any MT4 file:</strong> the <strong>account-type</strong> file with columns "
        "<code>account</code> and <code>account_type</code>. Without it the run stops before upload. "
        "Each MT4 row's Login is matched to <code>account</code>: <code>CentAccount</code> (any case or spacing, "
        "e.g. <code>Cent Account</code>, <code>cent_account</code>) is USC and its amounts are divided by 100 to USD. "
        "Classic, Plus, Pro, and every other type are USD.",
        "A Login missing from the account-type file is treated as USD and listed on the Verifikasi sheet. "
        "There is no USD/USC choice per file; a currency column in the MT4 file is only a cross-check.",
        "Each MT5 file must contain Deal, Login, Time, Type, Entry, Symbol, Volume, Commission, Fee, Swap, Profit and Currency. Group and Country are ignored.",
        "Optional: one <strong>Client Equity FX</strong> / rate file in <code>.xlsx</code> "
        "or <code>.csv</code> format with columns <code>date</code>, <code>currency</code>, "
        "<code>rate</code>. Any sheet name is fine and headers are not case-sensitive.",
        f"There is a <strong>{chunked.teks_batas()} total upload limit</strong> and a "
        f"<strong>{chunked.MAX_FILES}-file maximum</strong> per run. Large batches are sent to the server "
        "in pieces automatically.",
    ],
    "langkah": [
        "MT5 <code>Entry = out</code> and <code>Entry = in</code> stay on separate sheets. "
        "Nothing is merged between them, and Commission is never copied from an opening row onto a closing row.",
        "An MT4 closed buy/sell becomes an out row at Close Time (Profit, Swap, Commission, and Taxes as Fee). "
        "Opening rows are not written. Agent is left out of Profit, Commission and Fee. "
        "Cancelled pending orders, balance/credit rows (CRM swap, deposit, withdrawal, transfer) and the footer are excluded, "
        "the same rule as an MT5 row whose Entry is not in or out. They are counted on Verifikasi.",
        "When an MT4 file is uploaded, its out rows join the same <code>Daily</code> and <code>Monthly Summary</code> sheets as MT5, "
        "with a Platform column (<code>MT5</code> / <code>MT4</code>). Rows are grouped by Platform and Login, "
        "so the same login number on both platforms stays two rows. An MT5-only upload does not add that column.",
        "MT4 USC profit, commission, swap and fee are divided by 100. Volume is not. "
        "That conversion does not use the Client Equity FX rate. MT5 USC still does: "
        "missing MT5 rates stay blank and are highlighted yellow.",
    ],
    "hasil": "One <code>.zip</code> containing two workbooks:",
    "sheets": [
        ("Deals - Daily.xlsx",
         "<code>Daily</code> (closing rows) and <code>Daily - In</code> (MT5 opening rows) on separate sheets. "
         "When an MT4 Raw Report was uploaded, MT4 closing rows are on <code>Daily</code> too and every row has a Platform column. "
         "Plus Verifikasi, with totals per platform and a combined total."),
        ("Deals - Monthly Summary.xlsx",
         "<code>Monthly Summary</code> and <code>Monthly Summary - In</code> on separate sheets, "
         "with the same Platform column when MT4 was uploaded. Plus Verifikasi."),
    ],
    "catatan": ["Review the <strong>Verifikasi</strong> sheet before using the totals. "
                "It counts MT4 rows matched in the account-type file, lists logins missing from it "
                "(treated as USD), flags any currency-column disagreement, and notes MT4 USC ÷100. "
                "Agent is listed there as information and is not part of Profit, Commission or Fee.",
                "After a successful run the uploaded file is deleted from the server. "
                "The file on your computer is not changed. Verifikasi records this.",
                CATATAN_ASLI],
}


def _copy_for(company, slug):
    """Hint and guide for a tool page. KVB Generate D&W and KVB Deal Segregator differ from DPM."""
    if company == "kvb" and slug == "dw":
        return KVB_HINT, KVB_DOC
    if company == "kvb" and slug == "segregate":
        return KVB_SEGREGATE_HINT, KVB_SEGREGATE_DOC
    return HINT.get(slug), DOC.get(slug)


@app.context_processor
def _versi():
    # Every template, the sign-in page included (auth.py renders through this app).
    return {"app_version": VERSI["version"], "app_build": VERSI["build"],
            "app_version_label": VERSI_LABEL}


@app.get("/version")
def version():
    """Running version as JSON. Behind the login hook like every other route."""
    return {"name": APP_NAME, "version": VERSI["version"], "build": VERSI["build"],
            "label": VERSI_LABEL}


@app.errorhandler(HTTPException)
def _halaman_error(exc):
    """Werkzeug's own error page with the version under it (403/404/405/500...).

    Status, headers, and the description stay the same; only a footer line is added
    so a screenshot of an error shows which version is running.
    """
    resp = exc.get_response()
    if exc.code is None or exc.code < 400 or exc.response is not None:
        return resp
    if (resp.mimetype or "") != "text/html":
        return resp
    resp.set_data(exc.get_body().rstrip() + f'\n<p><small>{escape(VERSI_LABEL)}</small></p>\n')
    return resp


@app.context_processor
def _nav():
    company = getattr(request, "company", "kvb" if request.path.startswith("/kvb") else "dpm")
    user = getattr(request, "user", None)
    return {"tools": _tools_for(company) or {}, "company": company,
            "company_urls": {"dpm": "/", "kvb": "/kvb"},
            "can_dpm": auth.has_company(user, "dpm"),
            "can_kvb": auth.has_company(user, "kvb"),
            "max_files": chunked.MAX_FILES}


@app.template_filter("waktu")
def _waktu(ts):
    """Render a unix timestamp as a short local date, or a dash when never set."""
    if not ts:
        return "—"
    return datetime.datetime.fromtimestamp(ts).strftime("%d %b %Y, %H:%M")


@app.get("/")
def index():
    request.company = "dpm"
    return render_template("index.html")


@app.get("/kvb")
def kvb_index():
    request.company = "kvb"
    return render_template("index.html")


@app.get("/tool/<slug>")
def tool(slug, company="dpm"):
    registry = _tools_for(company)
    if slug not in registry or not registry[slug][3]:
        abort(404)
    request.company = company
    pilihan = _pilihan_bulan() if slug == "dw" else {}
    hint, doc = _copy_for(company, slug)
    return render_template("tool.html", slug=slug, tool=registry[slug],
                           hint=hint, doc=doc, **pilihan)


@app.get("/kvb/tool/<slug>")
def kvb_tool(slug):
    return tool(slug, "kvb")


@app.post("/tool/<slug>")
def run(slug, company="dpm"):
    registry = _tools_for(company)
    if slug not in registry or not registry[slug][3]:
        abort(404)
    request.company = company
    pilihan = _pilihan_bulan() if slug == "dw" else {}
    # Keep what the user picked, so a validation error does not reset the form.
    if (request.form.get("bulan") or "").isdigit():
        pilihan["bulan_default"] = int(request.form["bulan"])
    if (request.form.get("tahun") or "").isdigit():
        pilihan["tahun_default"] = int(request.form["tahun"])

    def _gagal(pesan):
        hint, doc = _copy_for(company, slug)
        return render_template("tool.html", slug=slug, tool=registry[slug],
                               hint=hint, doc=doc,
                               error=pesan, **pilihan), 400

    uploads = [u for u in request.files.getlist("file") if u and u.filename]
    # Large batches arrive in pieces via /upload/chunk and are already on disk by
    # the time this runs -- Cloudflare caps a single request body at 100 MB, so a
    # 900 MB run can only reach us that way. Both paths converge on the same list.
    upload_id = (request.form.get("upload_id") or "").strip()
    staged = chunked.staged_files(JOBS, upload_id) if upload_id else None
    if upload_id and not staged:
        return _gagal("The upload did not finish, or it expired. Please choose the files again.")
    pemilik = cleanup.upload_owner(JOBS, upload_id) if staged else None
    if pemilik and pemilik != request.user["username"]:
        return _gagal("This upload belongs to another account. Please choose the files again.")

    if slug in ("segregate", "segregated"):
        if not uploads and not staged:
            return _gagal("Please choose at least one .csv, .xlsx, or .xlsm file first.")
        # One count for every upload field: "file" (MT5 + MT4 Deals and the FX
        # workbook), files already staged by /upload/chunk, and the KVB
        # "mt4_accounts" account-type file.
        akun = request.files.get("mt4_accounts") if company == "kvb" else None
        n_akun = 1 if akun and akun.filename else 0
        if len(uploads) + len(staged or []) + n_akun > chunked.MAX_FILES:
            return _gagal(f"Select no more than {chunked.MAX_FILES} files per run.")
        empty = []
        for upload in uploads:
            pos = upload.stream.tell()
            upload.stream.seek(0, os.SEEK_END)
            size = upload.stream.tell()
            upload.stream.seek(pos)
            if size == 0:
                empty.append(upload.filename)
        empty += [original for path, original in (staged or []) if path.stat().st_size == 0]
        if empty:
            return _gagal("Empty files cannot be processed: " + ", ".join(empty))
        names = [u.filename for u in uploads] + [orig for _p, orig in (staged or [])]
        invalid = [n for n in names if not n.lower().endswith((".csv", ".xlsx", ".xlsm"))]
        if invalid:
            return _gagal("These files are not .csv, .xlsx, or .xlsm and were rejected: "
                          f"{', '.join(invalid)}")
        if slug == "segregated" and len(names) < 2:
            # The engine detects the FX file by its header and stops without it;
            # one file can never be both, so say so before anything is moved.
            return _gagal("Select the daily segregated files and the month's Client Equity "
                          "FX file. The Client Equity FX file is required.")
        if company == "kvb" and slug == "segregate":
            _akun_upload = request.files.get("mt4_accounts")
            if _akun_upload and _akun_upload.filename:
                if not _akun_upload.filename.lower().endswith((".csv", ".xlsx", ".xlsm")):
                    return _gagal("The MT4 account-type file must be .csv, .xlsx, or .xlsm.")
                pos = _akun_upload.stream.tell()
                _akun_upload.stream.seek(0, os.SEEK_END)
                if _akun_upload.stream.tell() == 0:
                    return _gagal("Empty files cannot be processed: " + _akun_upload.filename)
                _akun_upload.stream.seek(pos)
        else:
            _akun_upload = None
        periode = saldo_jw = saldo_ch = None
    elif slug == "pl-desk":
        names = [u.filename for u in uploads] + [orig for _p, orig in (staged or [])]
        if len(names) != 2:
            return _gagal("Upload two files: the FinanceOS settlement (PDF or converted Excel) "
                          "and the Metabase workbook.")
        invalid = [n for n in names if not n.lower().endswith((".pdf", ".xlsx", ".xlsm"))]
        if invalid:
            return _gagal("PL by Desk accepts a .pdf and an .xlsx workbook. These files were "
                          f"rejected: {', '.join(invalid)}")
        empty = []
        for upload in uploads:
            pos = upload.stream.tell()
            upload.stream.seek(0, os.SEEK_END)
            size = upload.stream.tell()
            upload.stream.seek(pos)
            if size == 0:
                empty.append(upload.filename)
        empty += [original for path, original in (staged or []) if path.stat().st_size == 0]
        if empty:
            return _gagal("Empty files cannot be processed: " + ", ".join(empty))
        periode = saldo_jw = saldo_ch = None
    else:
        first = uploads[0].filename if uploads else (staged[0][1] if staged else None)
        if not first or not first.lower().endswith((".xlsx", ".xlsm")):
            return _gagal("Please choose an .xlsx file first.")
        periode, salah = _baca_periode(request.form)
        if salah:
            return _gagal(salah)
        saldo_jw = (request.form.get("saldo_jw") or "").strip()
        pilihan["saldo_default"] = saldo_jw
        if saldo_jw:
            try:
                ok = (len(saldo_jw) <= SALDO_MAKS_PANJANG and angka.parse_number(
                    saldo_jw, ambiguous=angka.THOUSANDS) is not None)
            except angka.NumberError:
                ok = False
            if not ok:
                return _gagal(f"The opening balance must be a number, for example 377233.36, "
                              f"377,233.36, 377.233,36 or 3.5E+09. Got {saldo_jw[:60]!r} "
                              f"({VERSI_LABEL}). Leave it empty to start from zero.")
        saldo_ch = (request.form.get("saldo_channel") or "").strip()
        pilihan["saldo_ch_default"] = saldo_ch
        if saldo_ch and not any(c.isdigit() for c in saldo_ch):
            return _gagal("The channel opening balances do not contain a single number. "
                          "Paste three columns: Payment Channel, Currency, Balance.")

    def _ukuran_stream(upload):
        pos = upload.stream.tell()
        upload.stream.seek(0, os.SEEK_END)
        ukuran = upload.stream.tell()
        upload.stream.seek(pos)
        return ukuran

    ukuran_staged = sum(path.stat().st_size for path, _original in (staged or []))
    ukuran_baru = sum(_ukuran_stream(upload) for upload in uploads)
    if (slug == "segregate" and company == "kvb" and _akun_upload
            and _akun_upload.filename):
        ukuran_baru += _ukuran_stream(_akun_upload)
    # Lock the new job before its directory exists, so a cleanup in another
    # worker can never mistake it for leftovers. Then free the old data of
    # every other finished/abandoned run, and only then check the disk.
    job_id = uuid.uuid4().hex
    g.job_lock = cleanup.acquire_job_lock(JOBS, job_id)
    with cleanup.admission(JOBS):
        freed = cleanup.clean(JOBS, keep_jobs={job_id},
                              keep_uploads={upload_id} if upload_id else ())
        # Before any rename. A failed check must not delete the staged upload:
        # that file is the only server copy, and copying it is not allowed.
        salah_ruang = chunked.cek_sebelum_proses(JOBS, ukuran_staged, ukuran_baru,
                                                 freed=freed)
        if salah_ruang:
            return _gagal(salah_ruang)
        job = JOBS / job_id
        job.mkdir()
    sources = []
    nama_asli = []

    def _tujuan(original, urutan):
        """Pick a collision-free name inside the job directory."""
        safe = secure_filename(original) or f"upload-{urutan}"
        if not Path(safe).suffix:
            safe += Path(original).suffix.lower()
        target = job / safe
        n = 2
        while target.exists():
            target = job / f"{Path(safe).stem}-{n}{Path(safe).suffix}"
            n += 1
        return target

    if staged:
        # Compare directories, not the file: overlayfs (Docker, the dev box)
        # reports a different st_dev for regular files than for directories on
        # the same mount, which made every staged run look cross-device. A real
        # cross-device rename is still caught by the EXDEV handler below.
        dev = job.stat().st_dev
        beda = [original for path, original in staged
                if path.parent.stat().st_dev != dev]
        if beda:
            shutil.rmtree(job, ignore_errors=True)
            return _gagal("The upload is on a different disk from the job folder, so it "
                          "cannot be moved into place. Copying it would keep a second "
                          "copy of the file, which this server does not have room for. "
                          "Nothing was started. The upload was not deleted.")

    urutan = 0
    for upload in uploads:
        urutan += 1
        target = _tujuan(upload.filename, urutan)
        upload.save(target)
        sources.append(target)
        nama_asli.append(upload.filename)
    pindah = []
    try:
        for path, original in (staged or []):
            urutan += 1
            target = _tujuan(original, urutan)
            # Same-filesystem rename. A cross-device move would copy, so it is refused.
            os.replace(path, target)
            pindah.append((target, path))
            sources.append(target)
            nama_asli.append(original)
    except OSError as exc:
        if exc.errno not in (errno.EXDEV, errno.ENOENT):
            raise
        for tujuan, asal in reversed(pindah):
            try:
                os.replace(tujuan, asal)
            except OSError:
                pass
        shutil.rmtree(job, ignore_errors=True)
        if exc.errno == errno.ENOENT:
            # The same account started a newer upload a moment ago, which replaced this one.
            return _gagal("This upload was replaced by a newer upload from your account "
                          "(another tab or a new attempt). Nothing was started here.")
        return _gagal("The upload could not be moved into the job folder without "
                      "copying it. Nothing was started. The upload was not deleted.")
    if staged:
        chunked.discard(JOBS, upload_id)
    if company == "kvb" and slug == "segregate" and not (_akun_upload and _akun_upload.filename):
        salah_mt4 = _mt4_tanpa_peta_akun(sources, nama_asli)
        if salah_mt4:
            shutil.rmtree(job, ignore_errors=True)
            return _gagal(salah_mt4)
    src = sources[0]
    if slug == "dw":
        dst = src.with_name(f"{src.stem}-hasil.xlsx")
    elif slug == "pl-desk":
        dst = job / "PL by Desk.xlsx"
    elif slug == "segregated":
        dst = job / "Segregated.xlsx"
    else:
        dst = job / "hasil.zip"

    # Asynchronous on purpose. A 38 MB workbook takes several minutes, and any
    # proxy in front of this app (Cloudflare caps at 100s) would kill a request
    # that stays open that long -- the user saw "Error 524". So the upload
    # returns immediately with a job id, the work continues in a background
    # thread, and the browser polls a tiny status page until the file is ready.
    # The month goes in the download name on purpose: two runs of the same
    # upload for different months would otherwise be indistinguishable on disk.
    if slug == "dw":
        download_name = f"{Path(nama_asli[0]).stem} - {_label_periode(periode)}-hasil.xlsx"
    elif slug == "pl-desk":
        download_name = "PL by Desk.xlsx"
    elif slug == "segregated":
        download_name = "Segregated.xlsx"       # the month is added when the run ends
    else:
        stems = [Path(n).stem for n in nama_asli]
        label = " - ".join(stems) if len(stems) <= 3 else f"{len(stems)} files"
        download_name = f"{label}-hasil.zip"
    _write_state(job_id, state="running", name=download_name,
                 started=time.time(), error=None, log=None, path=None, slug=slug,
                 company=company, version=VERSI_LABEL,
                 user=request.user["username"],      # names the account in queue notices
                 owner_pid=os.getpid(),
                 periode=periode, saldo_jw=saldo_jw or None,
                 saldo_ch_baris=(len([x for x in saldo_ch.splitlines() if x.strip()])
                                 if saldo_ch else None))
    mt4_accounts_path = None
    if company == "kvb" and slug == "segregate" and _akun_upload and _akun_upload.filename:
        urutan += 1
        target = _tujuan(_akun_upload.filename, urutan)
        _akun_upload.save(str(target))
        mt4_accounts_path = target
    # The request returns now. gunicorn's 1800s timeout only kills a worker
    # that is stuck inside a request and cannot heartbeat. This thread, and
    # the subprocess it starts, keep running after the 303. There is no
    # subprocess timeout. Status polls are short GETs, so a run of a few hours
    # is not cut off. The staged file was renamed on this same volume.
    lock_fh = g.pop("job_lock", None)          # the thread owns the lock now
    threading.Thread(target=_process_job,
                     args=(job_id, company, slug, job,
                           sources if slug in ("segregate", "pl-desk", "segregated") else src,
                           dst, periode, saldo_jw, saldo_ch, mt4_accounts_path,
                           lock_fh),
                     daemon=True).start()
    endpoint = "kvb_job_status" if company == "kvb" else "job_status"
    return redirect(url_for(endpoint, job_id=job_id), code=303)


@app.post("/kvb/tool/<slug>")
def kvb_run(slug):
    return run(slug, "kvb")


def _cleanup_completed_job(job, result):
    """Delete source/intermediate data while retaining the downloadable result."""
    result = result.resolve()
    for entry in job.iterdir():
        if entry.resolve() == result:
            continue
        if entry.is_dir():
            shutil.rmtree(entry, ignore_errors=True)
        else:
            entry.unlink(missing_ok=True)


def _mt4_tanpa_peta_akun(paths, names):
    """STOP text when an MT4 Raw Report is uploaded without the account-type file.

    The account-type file may also arrive in the main file field; it is
    recognised from its account / account_type header. Unreadable files are
    left for the engine to report.
    """
    import deal_segregator as ds
    mt4, peta = [], False
    for path, name in zip(paths, names):
        try:
            if ds._adalah_peta_akun(path):
                peta = True
            elif ds._info_mt4(path):
                mt4.append(name)
        except (SystemExit, Exception):
            continue
    if mt4 and not peta:
        return ("STOP: the MT4 account-type file is required. MT4 file(s) uploaded: "
                f"{', '.join(mt4)}. Add the account-type file (columns account, "
                "account_type) and run again. CentAccount is USC (÷100); every other "
                "type is USD.")
    return None


def _pisah_traceback(stderr):
    """User-visible text, and the traceback kept for the server log."""
    pengguna, jejak = [], []
    di_jejak = False
    for baris in (stderr or "").splitlines():
        if baris.startswith("Traceback (most recent call last):"):
            di_jejak = True
        if di_jejak:
            jejak.append(baris)
        else:
            pengguna.append(baris)
    return "\n".join(pengguna), "\n".join(jejak)


def _pesan_gagal_job(stderr, stdout):
    """The line a user should read. A STOP line wins over a traceback."""
    for sumber in (stderr, stdout):
        for baris in (sumber or "").splitlines():
            teks = baris.strip()
            if teks.startswith("STOP:"):
                return teks
    for sumber in (stderr, stdout):
        for baris in (sumber or "").splitlines():
            teks = baris.strip()
            if not teks or teks.startswith("Traceback") or teks.startswith("File \""):
                continue
            if teks.startswith("During handling") or teks[:1] in " ^~":
                continue
            return teks
    return "STOP: the run failed."


def _process_job(job_id, company, slug, job, src, dst, periode=None, saldo_jw=None,
                 saldo_ch=None, mt4_accounts=None, lock_fh=None):
    """Run the pipeline in the background and record the outcome.

    lock_fh is the job's flock; while it is held no cleanup touches this job.
    """
    try:
        _jalankan_job(job_id, company, slug, job, src, dst, periode, saldo_jw,
                      saldo_ch, mt4_accounts)
    finally:
        cleanup.release_job_lock(lock_fh)


def _jalankan_job(job_id, company, slug, job, src, dst, periode=None, saldo_jw=None,
                  saldo_ch=None, mt4_accounts=None):
    # Success: inputs and the deal index are deleted, only the result stays.
    # Failure or STOP: the whole job directory (inputs, index, partial output)
    # is deleted; the staged upload was already moved into it.
    # Tabel saldo channel yang ditempel disimpan sebagai file di folder job -- ikut
    # terhapus bersama job-nya, jadi tidak ada sisa data keuangan yang menetap.
    f_ch = None
    if saldo_ch:
        f_ch = job / "saldo-channel.txt"
        f_ch.write_text(saldo_ch, encoding="utf-8")
    langkah = _tools_for(company)[slug][2]
    if isinstance(langkah, str):
        langkah = (langkah,)
    log = []
    masuk = src
    try:
        for i, script in enumerate(langkah):
            keluar = dst if i == len(langkah) - 1 else job / f"step{i}.xlsx"
            inputs = masuk if isinstance(masuk, (list, tuple)) else [masuk]
            perintah = [sys.executable, str(BASE / script), *(str(p) for p in inputs),
                        "-o", str(keluar)]
            flag = PERIODE_ARG.get(script)
            if periode and flag:
                perintah += [flag, periode]
            flag_saldo = SALDO_ARG.get(script)
            if saldo_jw and flag_saldo:
                # '--flag=value' so a value starting with '-' ('-Rp 1.000')
                # is never mistaken for another option.
                perintah += [f"{flag_saldo}={saldo_jw}"]
            if company == "kvb" and KVB_ARG.get(script):
                tambahan = KVB_ARG[script]
                perintah.extend(tambahan if isinstance(tambahan, (list, tuple)) else [tambahan])
            if script in ("deal_segregator.py", "dpm_segregated.py"):
                # The workbook says so on Verifikasi. The CLI does not pass this,
                # so a command-line run leaves the source files where they are.
                perintah.append("--delete-input")
            if company == "kvb" and script == "deal_segregator.py":
                perintah.append("--allow-mt4")
                if mt4_accounts:
                    perintah += ["--mt4-accounts", str(mt4_accounts)]
            flag_ch = CHANNEL_ARG.get(script)
            if f_ch and flag_ch:
                perintah += [flag_ch, str(f_ch)]
            # No timeout. A 10 GB Deal Segregator run can take a few hours,
            # and the worker has already finished the HTTP request.
            r = subprocess.run(perintah, capture_output=True,
                               text=True, cwd=BASE)
            pengguna, jejak = _pisah_traceback(r.stderr or "")
            if jejak:
                print(f"job {job_id} {script}\n{jejak}", file=sys.stderr)
            log.append(f"$ {script}\n{r.stdout}{pengguna}".rstrip())
            if r.returncode != 0 or not keluar.exists():
                judul = _pesan_gagal_job(r.stderr or "", r.stdout or "")
                _write_state(job_id, state="failed", error=judul or None,
                             log="\n\n".join(log).strip() or judul or "Failed with no message.")
                shutil.rmtree(job, ignore_errors=True)
                return
            masuk = keluar
        _cleanup_completed_job(job, dst)
        done = {"state": "done", "path": str(dst), "log": "\n\n".join(log).strip()}
        if slug == "pl-desk":
            matched = re.search(r"(?m)^PERIOD:\s*(\d{4}-\d{2})\s*$", done["log"])
            if matched:
                done["name"] = f"PL by Desk - {_label_periode(matched.group(1))}.xlsx"
        if slug == "segregated":
            matched = re.search(r"(?m)^PERIOD:\s*(\d{4}-\d{2})\s*$", done["log"])
            if matched:
                done["name"] = f"Segregated - {_label_periode(matched.group(1))}.xlsx"
        _write_state(job_id, **done)
    except Exception as exc:                       # keep the worker from dying silently
        traceback.print_exc(file=sys.stderr)
        pesan = f"STOP: {type(exc).__name__}: {exc}"
        _write_state(job_id, state="failed", error=pesan,
                     log="\n\n".join(log).strip() or pesan)
        shutil.rmtree(job, ignore_errors=True)


@app.get("/kvb/job/<job_id>", endpoint="kvb_job_status")
@app.get("/job/<job_id>")
def job_status(job_id):
    info = _read_state(job_id)
    if not info:
        abort(404)
    company = info.get("company", "dpm")
    denied = _require_company(company)
    if denied is not None:
        return denied
    request.company = company
    registry = _tools_for(company)
    slug = info["slug"]
    if info["state"] == "failed":
        # Keep the directory: a refresh on the error page must still show the
        # error, not a bare 404. The sweeper removes it later.
        pilihan = _pilihan_bulan() if slug == "dw" else {}
        hint, doc = _copy_for(company, slug)
        return render_template("tool.html", slug=slug, tool=registry[slug],
                               hint=hint, doc=doc,
                               error=info["error"],
                               log=info["log"] or "Failed with no message.",
                               job_version=info.get("version"),
                               **pilihan), 422
    elapsed = int(time.time() - info["started"])
    return render_template("job.html", job_id=job_id, info=info, elapsed=elapsed,
                           tool=registry[slug])


@app.get("/kvb/job/<job_id>/download", endpoint="kvb_job_download")
@app.get("/job/<job_id>/download")
def job_download(job_id):
    info = _read_state(job_id)
    if not info:
        abort(404)
    company = info.get("company", "dpm")
    # A download is not a page. Redirecting it could point at the other
    # company's file, so a browser still gets 403 and an empty body.
    _require_company(company, browser_redirect=False)
    request.company = company
    registry = _tools_for(company)
    if info["state"] == "cleared":
        return _halaman_dibersihkan(job_id, info, registry)
    if info["state"] == "collected":
        # Already downloaded and wiped. Say so plainly instead of a bare 404 --
        # a second click or a browser retry lands here and must not look broken.
        return render_template("job.html", job_id=job_id, info=info, elapsed=0,
                               tool=registry[info["slug"]]), 410
    if info["state"] != "done" or not info["path"]:
        abort(404)
    job = _job_dir(job_id)
    path = Path(info["path"])
    claimed = JOBS / f".{job_id}.download"
    if not job.is_dir() and not claimed.is_dir():
        # A newer run's cleanup removed it to free disk.
        return _halaman_dibersihkan(job_id, info, registry)
    # Holding the job lock while streaming keeps a cleanup started by another
    # user's new run from deleting the file mid-download.
    kunci = cleanup.acquire_job_lock(JOBS, job_id)
    if kunci is None:
        return "Download already in progress.", 409
    try:
        os.rename(job, claimed)                 # atomic claim across Gunicorn workers
    except OSError:
        cleanup.release_job_lock(kunci)
        latest = _read_state(job_id)
        if latest and latest["state"] == "collected":
            return render_template("job.html", job_id=job_id, info=latest, elapsed=0,
                                   tool=registry[latest["slug"]]), 410
        if latest and latest["state"] == "cleared" or not claimed.is_dir():
            return _halaman_dibersihkan(job_id, latest or info, registry)
        return "Download already in progress.", 409

    result = claimed / path.relative_to(job)
    if not result.is_file():
        shutil.rmtree(claimed, ignore_errors=True)
        cleanup.release_job_lock(kunci)
        return _halaman_dibersihkan(job_id, info, registry)
    mimetype = ("application/zip" if info["slug"] == "segregate" else
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response = send_file(result, as_attachment=True,
                         download_name=info["name"], mimetype=mimetype)
    stream = response.response

    def cleanup_stream():
        completed = False
        try:
            yield from stream
            completed = True
        finally:
            close = getattr(stream, "close", None)
            if close:
                close()
            if completed:
                _write_state(job_id, state="collected", path=None, collected=time.time())
                shutil.rmtree(claimed, ignore_errors=True)
            else:
                # A dropped browser/VPN connection must not destroy a finished
                # financial artifact. Restore it so the same URL can be retried.
                try:
                    os.rename(claimed, job)
                except OSError:
                    pass
            cleanup.release_job_lock(kunci)

    response.response = cleanup_stream()
    return response


def _halaman_dibersihkan(job_id, info, registry):
    """Friendly 410 when a newer run's cleanup removed this result."""
    info = dict(info)
    if info.get("state") != "cleared":
        info = _write_state(job_id, state="cleared", path=None, cleared=time.time())
    return render_template("job.html", job_id=job_id, info=info, elapsed=0,
                           tool=_tools_for(info.get("company", "dpm"))[info["slug"]]), 410


def _pid_alive(pid):
    try:
        os.kill(int(pid), 0)
        return True
    except (OSError, TypeError, ValueError):
        return False


def _sweep_old_jobs(max_age_hours=1):
    """Remove stale job directories and their state markers.

    Downloaded jobs delete themselves immediately; this only catches jobs that
    were abandoned (browser closed mid-run) or that failed. Runs on each upload,
    so no scheduler is needed.
    """
    cutoff = time.time() - max_age_hours * 3600
    for entry in JOBS.iterdir():
        try:
            if entry.name == "staging":
                continue
            job_id = entry.name if entry.is_dir() else entry.stem
            info = _read_state(job_id)
            if (info and info.get("state") == "running" and
                    _pid_alive(info.get("owner_pid"))):
                continue
            if entry.stat().st_mtime >= cutoff:
                continue
            if entry.is_dir():
                shutil.rmtree(entry, ignore_errors=True)
            else:
                entry.unlink(missing_ok=True)
        except OSError:
            pass


@app.get("/template")
def template_file():
    f = BASE / "Template D&W.xlsx"
    return send_file(f, as_attachment=True) if f.is_file() else abort(404)


@app.get("/kvb/template")
def kvb_template_file():
    f = BASE / "Template KVB Plus.xlsx"
    return send_file(f, as_attachment=True) if f.is_file() else abort(404)


def _kirim_template_segregator(nama):
    from segregator_templates import NAMA_BERKAS, bytes_template
    if nama not in NAMA_BERKAS:
        abort(404)
    return send_file(
        bytes_template(nama),
        as_attachment=True,
        download_name=NAMA_BERKAS[nama],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/tool/segregate/template/<nama>")
def template_segregator_dpm(nama):
    if nama not in ("mt5", "fx"):
        abort(404)
    return _kirim_template_segregator(nama)


@app.get("/tool/segregated/template/<nama>")
def template_segregated_dpm(nama):
    """DPM Segregated: the daily layout, and the same Client Equity FX template."""
    if nama not in ("segregated", "fx"):
        abort(404)
    return _kirim_template_segregator(nama)


@app.get("/kvb/tool/segregate/template/<nama>")
def template_segregator_kvb(nama):
    if nama not in ("mt5", "mt4", "accounts", "fx"):
        abort(404)
    return _kirim_template_segregator(nama)


if __name__ == "__main__":
    # BAHAYA KALAU DEPLOY: debug=True membuka debugger Werkzeug, dan siapa pun yang
    # memicu error bisa menjalankan kode Python di server. Sekarang MATI secara
    # default; nyalakan hanya saat mengembangkan di laptop sendiri:
    #     DW_DEBUG=1 python3 app.py
    #
    # app.run() itu server pengembangan, bukan untuk produksi. Untuk deploy pakai
    # WSGI yang benar, mis.:
    #     python3 -m pip install --user waitress
    #     python3 -m waitress --host 127.0.0.1 --port 5000 app:app
    # lalu taruh di belakang nginx/Caddy kalau perlu diakses dari luar.
    debug = os.environ.get("DW_DEBUG") == "1"
    app.run(host=os.environ.get("DW_HOST", "127.0.0.1"),
            port=int(os.environ.get("DW_PORT", "5000")),
            debug=debug)
