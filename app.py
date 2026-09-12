#!/usr/bin/env python3
"""dw-calculator web UI. Landing page + menu generator.

Jalankan:  python3 app.py     -> http://127.0.0.1:5000
Menu baru: tambahkan entry di TOOLS. Tidak perlu ubah kode lain.
"""
import datetime
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
import uuid
from pathlib import Path

from flask import (Flask, abort, redirect, render_template, request, send_file,
                   url_for)
from werkzeug.utils import secure_filename

BASE = Path(__file__).resolve().parent
JOBS = (Path(tempfile.gettempdir()) / "dw-calculator-jobs").resolve()
JOBS.mkdir(exist_ok=True)


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
    "segregate": ("Deal Segregator", "Combine one or more MT5 Deals History files into "
                  "daily and monthly summaries with verification totals.",
                  ("deal_segregator.py",), True),
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


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
SALDO_POLA = re.compile(r"^-?[\d.,\s]{1,24}$")
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
    "segregate": "Choose one or more MT5 Deals History files. CSV, XLSX and XLSM are accepted.",
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
        "ringkas": "Upload one or more MT5 Deals History exports. They are combined, filtered, "
                   "deduplicated and returned as one workbook.",
        "siapkan": [
            "One or more MT5 Deals History files in <code>.csv</code>, <code>.xlsx</code> or <code>.xlsm</code> format.",
            "Each file must contain Deal, Login, Group, Country, Time, Type, Entry, Symbol, Volume, Commission, Fee, Swap, Profit and Currency columns.",
        ],
        "langkah": [
            "Rows from every file are pooled. Only <code>Entry = out</code> rows are kept.",
            "Repeated non-empty Deal IDs are removed, then Daily and Monthly Summary totals are grouped by account, country, desk, period, type, symbol and currency.",
            "Country is copied exactly as supplied. No lookup or enrichment is performed.",
        ],
        "hasil": "One <code>.xlsx</code> workbook with three sheets:",
        "sheets": [("Daily", "Daily aggregates and financial totals."),
                   ("Monthly Summary", "Monthly aggregates and financial totals."),
                   ("Verifikasi", "Input, filtering, duplicate and output checks.")],
        "catatan": ["Review the <strong>Verifikasi</strong> sheet before using the totals.", CATATAN_ASLI],
    },
}


@app.context_processor
def _nav():
    return {"tools": TOOLS}


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/tool/<slug>")
def tool(slug):
    if slug not in TOOLS or not TOOLS[slug][3]:
        abort(404)
    pilihan = _pilihan_bulan() if slug == "dw" else {}
    return render_template("tool.html", slug=slug, tool=TOOLS[slug],
                           hint=HINT.get(slug), doc=DOC.get(slug), **pilihan)


@app.post("/tool/<slug>")
def run(slug):
    if slug not in TOOLS or not TOOLS[slug][3]:
        abort(404)
    pilihan = _pilihan_bulan() if slug == "dw" else {}
    # Keep what the user picked, so a validation error does not reset the form.
    if (request.form.get("bulan") or "").isdigit():
        pilihan["bulan_default"] = int(request.form["bulan"])
    if (request.form.get("tahun") or "").isdigit():
        pilihan["tahun_default"] = int(request.form["tahun"])

    def _gagal(pesan):
        return render_template("tool.html", slug=slug, tool=TOOLS[slug],
                               hint=HINT.get(slug), doc=DOC.get(slug),
                               error=pesan, **pilihan), 400

    uploads = [u for u in request.files.getlist("file") if u and u.filename]
    if slug == "segregate":
        if not uploads:
            return _gagal("Please choose at least one .csv or .xlsx file first.")
        invalid = [u.filename for u in uploads
                   if not u.filename.lower().endswith((".csv", ".xlsx", ".xlsm"))]
        if invalid:
            return _gagal(f"These files are not .csv/.xlsx and were rejected: {', '.join(invalid)}")
        periode = saldo_jw = None
    else:
        up = uploads[0] if uploads else None
        if not up or not up.filename.lower().endswith((".xlsx", ".xlsm")):
            return _gagal("Please choose an .xlsx file first.")
        periode, salah = _baca_periode(request.form)
        if salah:
            return _gagal(salah)
        saldo_jw = (request.form.get("saldo_jw") or "").strip()
        pilihan["saldo_default"] = saldo_jw
        if saldo_jw and not SALDO_POLA.match(saldo_jw):
            return _gagal("The opening balance must be a number, for example 377233.36 "
                          "or 377,233.36. Leave it empty to start from zero.")

    job = JOBS / uuid.uuid4().hex
    job.mkdir()
    _sweep_old_jobs()
    sources = []
    for index, upload in enumerate(uploads, start=1):
        safe = secure_filename(upload.filename) or f"upload-{index}"
        if not Path(safe).suffix:
            safe += Path(upload.filename).suffix.lower()
        src = job / safe
        suffix = 2
        while src.exists():
            src = job / f"{Path(safe).stem}-{suffix}{Path(safe).suffix}"
            suffix += 1
        upload.save(src)
        sources.append(src)
    src = sources[0]
    dst = src.with_name(f"{src.stem}-hasil.xlsx") if slug == "dw" else job / "hasil.xlsx"

    # Asynchronous on purpose. A 38 MB workbook takes several minutes, and any
    # proxy in front of this app (Cloudflare caps at 100s) would kill a request
    # that stays open that long -- the user saw "Error 524". So the upload
    # returns immediately with a job id, the work continues in a background
    # thread, and the browser polls a tiny status page until the file is ready.
    job_id = job.name
    # The month goes in the download name on purpose: two runs of the same
    # upload for different months would otherwise be indistinguishable on disk.
    if slug == "dw":
        download_name = f"{Path(uploads[0].filename).stem} - {_label_periode(periode)}-hasil.xlsx"
    else:
        stems = [Path(upload.filename).stem for upload in uploads]
        label = " - ".join(stems) if len(stems) <= 3 else f"{len(stems)} files"
        download_name = f"{label}-hasil.xlsx"
    _write_state(job_id, state="running", name=download_name,
                 started=time.time(), error=None, log=None, path=None, slug=slug,
                 periode=periode, saldo_jw=saldo_jw or None)
    threading.Thread(target=_process_job,
                     args=(job_id, slug, job, sources if slug == "segregate" else src,
                           dst, periode, saldo_jw),
                     daemon=True).start()
    return redirect(url_for("job_status", job_id=job_id))


def _process_job(job_id, slug, job, src, dst, periode=None, saldo_jw=None):
    """Run the pipeline in the background and record the outcome."""
    langkah = TOOLS[slug][2]
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
                perintah += [flag_saldo, saldo_jw]
            r = subprocess.run(perintah, capture_output=True,
                               text=True, cwd=BASE)
            log.append(f"$ {script}\n{r.stdout}{r.stderr}".rstrip())
            if r.returncode != 0 or not keluar.exists():
                pesan = (r.stderr or r.stdout).strip().splitlines()
                judul = next((x.strip() for x in pesan if x.strip()), "") if pesan else ""
                _write_state(job_id, state="failed", error=judul or None,
                             log="\n\n".join(log).strip() or "Failed with no message.")
                shutil.rmtree(job, ignore_errors=True)
                return
            masuk = keluar
        _write_state(job_id, state="done", path=str(dst),
                     log="\n\n".join(log).strip())
    except Exception as exc:                       # keep the worker from dying silently
        _write_state(job_id, state="failed", error=f"{type(exc).__name__}: {exc}",
                     log="\n\n".join(log).strip() or None)
        shutil.rmtree(job, ignore_errors=True)


@app.get("/job/<job_id>")
def job_status(job_id):
    info = _read_state(job_id)
    if not info:
        abort(404)
    slug = info["slug"]
    if info["state"] == "failed":
        # Keep the directory: a refresh on the error page must still show the
        # error, not a bare 404. The sweeper removes it later.
        pilihan = _pilihan_bulan() if slug == "dw" else {}
        return render_template("tool.html", slug=slug, tool=TOOLS[slug],
                               hint=HINT.get(slug), doc=DOC.get(slug),
                               error=info["error"],
                               log=info["log"] or "Failed with no message.",
                               **pilihan), 422
    elapsed = int(time.time() - info["started"])
    return render_template("job.html", job_id=job_id, info=info, elapsed=elapsed,
                           tool=TOOLS[slug])


@app.get("/job/<job_id>/download")
def job_download(job_id):
    info = _read_state(job_id)
    if not info:
        abort(404)
    if info["state"] == "collected":
        # Already downloaded and wiped. Say so plainly instead of a bare 404 --
        # a second click or a browser retry lands here and must not look broken.
        return render_template("job.html", job_id=job_id, info=info, elapsed=0,
                               tool=TOOLS[info["slug"]]), 410
    if info["state"] != "done" or not info["path"]:
        abort(404)
    path = Path(info["path"])
    if not path.is_file():
        abort(404)

    # Results can be 50 MB+, so nothing is kept on disk after delivery. Read into
    # memory first, then delete the job directory and mark the job collected, so
    # the file is gone the moment the response is handed to the client.
    data = path.read_bytes()
    shutil.rmtree(_job_dir(job_id), ignore_errors=True)
    _write_state(job_id, state="collected", path=None, collected=time.time())
    return send_file(io.BytesIO(data), as_attachment=True,
                     download_name=info["name"],
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _sweep_old_jobs(max_age_hours=6):
    """Remove stale job directories and their state markers.

    Downloaded jobs delete themselves immediately; this only catches jobs that
    were abandoned (browser closed mid-run) or that failed. Runs on each upload,
    so no scheduler is needed.
    """
    cutoff = time.time() - max_age_hours * 3600
    for entry in JOBS.iterdir():
        try:
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
