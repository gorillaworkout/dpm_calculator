#!/usr/bin/env python3
"""Deal Segregator -- segregasi export "Deals History" (MT5) per Login.

Jalankan:
    python3 deal_segregator.py "Juni.csv" "Client Equity - FX.xlsx" -o hasil.zip

Bisa menerima SATU ATAU LEBIH file sumber sekaligus (.csv atau .xlsx/.xlsm) --
mis. upload Juni dan Juli sebagai dua file terpisah, hasilnya DIGABUNG jadi
satu laporan. Cocok untuk laporan yang mencakup beberapa bulan tapi
diekspor per bulan dari MT5.

Tahap 1 (dikonfirmasi user 11 Sep 2026, direvisi 12, 12, 24 Sep & 6 Okt 2026):
  - Baris "out" (yang MENUTUP posisi -> punya Profit realized) dan baris "in"
    (yang MEMBUKA posisi) sama-sama diproses, tapi di SHEET TERPISAH:
    "Daily"/"Monthly Summary" = out saja, "Daily - In"/"Monthly Summary - In" =
    in saja. Layout kolom kedua sheet SAMA. Tidak ada penggabungan antar
    keduanya (keputusan user 6 Okt 2026, opsi 2 + 'Opsi A': tim Malaysia mau
    bisa mengecek bahwa in DAN out sama-sama masuk). Commission/Fee/Swap baris
    "in" tinggal di sheet "in" apa adanya dari MT5, TIDAK ditempel ke baris
    "out" (cara lama 24 Sep, dihapus supaya komisi tidak terhitung dua kali).
  - File .xlsx boleh punya beberapa sheet (mis. sheet IN dan sheet OUT): SEMUA
    sheet yang punya kolom Deals History dibaca, bukan cuma sheet aktif.
  - "Group" dan "Country" diabaikan sepenuhnya (keputusan user 24 Sep 2026).
    Keduanya tidak wajib, tidak ditulis ke output, dan tidak menjadi grouping.
  - "Position" tidak wajib (6 Okt 2026): tidak dipakai untuk menggabungkan in/out.
  - "Time" dipangkas jadi TANGGAL saja (jam/menit/detik dibuang).
  - Kalau ada beberapa file, baris dengan "Deal" ID yang SAMA (dari file
    manapun) hanya dihitung SEKALI -- jaga-jaga kalau dua file yang diupload
    ternyata tumpang tindih tanggalnya.
  - Hasilnya berupa .zip berisi Deals - Daily.xlsx dan
    Deals - Monthly Summary.xlsx; masing-masing punya sheet out, sheet in, dan
    Verifikasi. Baris dengan kunci yang sama pada level masing-masing
    digabung, kolom numerik dijumlah, kolom "Deals" mencatat berapa baris asli
    yang tergabung.

Skalabilitas (tetap dari perbaikan OOM di main, bukan dari salinan 6 Okt):
  - ID Deal yang sudah terpakai disimpan di SQLite di disk, bukan set Python,
    supaya run multi-GB tidak menumpuk ratusan juta ID di RAM.
  - Workbook ditulis write-only dan di-stream ke file sementara sebelum masuk
    zip, supaya ukuran output tidak ikut membesar di memori.

File sumber MT5 "Deals History" biasanya diekspor sebagai UTF-16LE
tab-delimited (bukan CSV koma biasa) -- encoding & pemisah kolom dideteksi
otomatis dari file, jadi export UTF-8/koma biasa dan file .xlsx juga tetap terbaca.

MT4 (KVB saja, flag --allow-mt4): file "Raw Report" dikenali dari baris judul
dan header, bukan dari nama file. Hanya baris buy/sell yang sudah ditutup
(Close Time) yang masuk, ke sheet Daily / Monthly Summary yang sama dengan MT5,
dengan kolom Platform. Login yang sama di MT4 dan MT5 tidak digabung. Tanpa
file MT4, sheet dan angka MT5 tidak berubah (kolom Platform tidak ditambahkan).
Mata uang MT4 tidak ada di file; pemanggil memilih USD atau USC per file
(default USD). Agent tidak masuk ke total manapun.
"""
import argparse
import csv
import datetime
import math
import re
import sqlite3
import tempfile
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

KOLOM_WAJIB = ["Deal", "Login", "Time", "Type", "Entry",
               "Symbol", "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]

KOLOM_HASIL = ["Login", "{PERIODE}", "Type", "Symbol", "Deals",
               "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]
KOLOM_USD = ["Commission (USD)", "Fee (USD)", "Swap (USD)", "Profit (USD)"]

# Ditulis hanya kalau ada file MT4. Upload MT5 saja tetap header lama.
KOLOM_HASIL_PLATFORM = ["Platform", "Login", "{PERIODE}", "Type", "Symbol", "Deals",
                        "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]

KOLOM_JUMLAH = ["Volume", "Commission", "Fee", "Swap", "Profit"]
KOLOM_TEKS = {"Platform", "Login", "Type", "Symbol", "Currency"}
KOLOM_UANG = ("Commission", "Fee", "Swap", "Profit")

# MT4 Manager "Raw Report". Judul + header, bukan nama file.
_MT4_JUDUL = re.compile(
    r"^Raw Report for ['\"][^'\"]*['\"] from (\d{4}\.\d{2}\.\d{2}) to (\d{4}\.\d{2}\.\d{2})\s*$",
    re.IGNORECASE,
)
_MT4_WAJIB = ["Deal", "Login", "Open Time", "Type", "Symbol", "Volume",
              "Close Time", "Agent", "Commission", "Taxes", "Swap", "Profit"]
_MT4_MATA_UANG = ("USD", "USC")

# Label Inggris untuk ringkasan yang DITAMPILKAN ke user (log job di web app).
# Kunci dict hasil tetap dipakai internal oleh pemanggil lain.
LABEL_RINGKASAN = {
    "file_masuk": "files read",
    "total": "source rows",
    "dipakai": "kept",
    "harian": "Daily rows",
    "bulanan": "Monthly Summary rows",
    "login_unik": "unique logins",
    "harian_in": "Daily 'in' rows",
    "bulanan_in": "Monthly Summary 'in' rows",
    "dipakai_in": "'in' rows kept",
}

ENTRY_DIPAKAI = ("out", "in")
FX_SHEET = "QUERY RESULT"
_DEAL_COMMIT_EVERY = 10000


# --------------------------------------------------------------------- baca file
def _deteksi_encoding(path: Path) -> str:
    """Deteksi UTF-16 (LE/BE) dari BOM, kalau tidak ada BOM anggap UTF-8."""
    with open(path, "rb") as f:
        head = f.read(4)
    if head.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16"
    if head.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    return "utf-8"


def _deteksi_pemisah(baris_pertama: str) -> str:
    return "\t" if baris_pertama.count("\t") >= baris_pertama.count(",") else ","


def _baca_csv(path: Path):
    enc = _deteksi_encoding(path)
    with open(path, encoding=enc, newline="") as f:
        awal = f.readline()
        pemisah = _deteksi_pemisah(awal)
    def rows():
        with open(path, encoding=enc, newline="") as f:
            r = csv.reader(f, delimiter=pemisah)
            next(r)
            for nomor, row in enumerate(r, start=2):
                if any((c or "").strip() for c in row):
                    yield nomor, row
    return next(csv.reader([awal], delimiter=pemisah)), rows(), None


def _baca_xlsx(path: Path):
    """Baca SEMUA sheet yang berisi kolom Deals History (bukan cuma sheet aktif).
    Sheet lain (tanpa kolom wajib) dilewati. Baris tiap sheet disusun ulang
    mengikuti urutan kolom sheet pertama yang cocok, jadi sheet dengan urutan
    kolom berbeda tetap terbaca benar. -> header, rows, close, info"""
    wb = None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        sheets, dilewati, header_pertama = [], [], []
        for ws in wb.worksheets:
            it = ws.iter_rows(values_only=True)
            first = next(it, None)
            header = [_teks(v) for v in first] if first else []
            if header and not header_pertama:
                header_pertama = header
            if all(k in header for k in KOLOM_WAJIB):
                sheets.append((ws.title, header, it))
            else:
                dilewati.append(ws.title)
    except Exception as exc:
        if wb is not None:
            wb.close()
        sys.exit(f"STOP: could not read the header in '{path.name}': {exc}")
    info = {"dibaca": {}, "dilewati": dilewati}
    if not sheets:
        wb.close()
        if not header_pertama:
            sys.exit(f"STOP: could not read the header in '{path.name}': the file has no header row")
        return header_pertama, iter(()), None, info
    kanon = sheets[0][1]

    def rows():
        try:
            for nama, header, it in sheets:
                posisi = [header.index(k) if k in header else None for k in kanon]
                n = 0
                for nomor, row in enumerate(it, start=2):
                    if any(v is not None and _teks(v) != "" for v in row):
                        n += 1
                        yield nomor, [row[i] if i is not None and i < len(row) else None
                                      for i in posisi]
                info["dibaca"][nama] = n
        finally:
            wb.close()
    return kanon, rows(), wb.close, info


def baca_file(path: Path):
    """Baca satu file export Deals History (.csv atau .xlsx/.xlsm). -> (idx, rows, info).
    info['dibaca'] = {nama_sheet: jumlah_baris} untuk .xlsx; baru terisi setelah
    rows habis dibaca. Kosong untuk .csv."""
    suffix = path.suffix.lower()
    info = {"dibaca": {}, "dilewati": []}
    if suffix in (".xlsx", ".xlsm"):
        header, rows, close, info = _baca_xlsx(path)
    else:
        header, rows, close = _baca_csv(path)
    idx = {name: i for i, name in enumerate(header)}
    hilang = [k for k in KOLOM_WAJIB if k not in idx]
    if hilang:
        if close:
            close()
        sys.exit(
            f"STOP: required columns are missing in '{path.name}': {hilang}\n"
            f"Header that was read ({len(header)} columns): {header}\n"
            "Make sure this file is a 'Deals History' export from MT5, not another file."
        )
    return idx, rows, info


# ------------------------------------------------------------------- normalisasi
def _teks(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()


def _norm(v) -> str:
    return _teks(v).upper()


def _adalah_file_fx(path: Path) -> bool:
    """Identify Client Equity FX by sheet and header, never by filename."""
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    wb = None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        sheet = next((s for s in wb.sheetnames if _norm(s) == FX_SHEET), None)
        if not sheet:
            return False
        header = [_teks(v) for v in next(wb[sheet].iter_rows(max_row=1, values_only=True), [])]
        return all(k in header for k in ("date", "Currency", "rate"))
    except Exception:
        return False
    finally:
        if wb is not None:
            wb.close()


def baca_fx(path: Path):
    """Read {(date, CURRENCY): rate}; blank rate means already USD (rate 1)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = next(s for s in wb.sheetnames if _norm(s) == FX_SHEET)
        it = wb[sheet].iter_rows(values_only=True)
        header = [_teks(v) for v in next(it)]
        idx = {name: i for i, name in enumerate(header)}
        result = {}
        for number, row in enumerate(it, start=2):
            if not any(v is not None and _teks(v) for v in row):
                continue
            date = tanggal_dari_waktu(row[idx["date"]], path, number)
            currency = _norm(row[idx["Currency"]])
            if not currency:
                sys.exit(f"STOP: empty Currency in FX file '{path.name}', row {number}")
            raw = row[idx["rate"]]
            if raw is None or _teks(raw) == "":
                # Only USD is already denominated in USD. A blank rate for any
                # other currency means conversion data is missing, not rate 1.
                rate = None if currency == "USD" else 0.0
            else:
                rate = _angka(raw, path, number, "rate")
                if rate <= 0:
                    sys.exit(f"STOP: FX rate must be positive in '{path.name}', row {number}: {rate!r}")
            key = (date, currency)
            if key in result:
                sys.exit(f"STOP: duplicate FX date/currency in '{path.name}', row {number}: "
                         f"{date} {currency}")
            result[key] = rate
        return result
    finally:
        wb.close()


def _angka(v, path: Path, nomor: int, kolom: str):
    teks = _teks(v)
    try:
        angka = float(teks)
    except ValueError:
        angka = math.nan
    if not teks or not math.isfinite(angka):
        sys.exit(f"STOP: invalid {kolom} value in '{path.name}', row {nomor}: {teks!r}")
    return angka


def desk_dari_group(group: str) -> str:
    g = (group or "").strip()
    if g.lower().startswith("real\\"):
        g = g[len("real\\"):]
    return g


def tanggal_dari_waktu(waktu, path=None, nomor=None, kolom="Time"):
    """Terima string MT5 ('2026.07.31 01:10:35.454') ATAU objek datetime/date
    (kalau dibaca dari .xlsx yang selnya sudah bertipe tanggal). -> date.
    Nilai kosong, format asing, dan tanggal mustahil ditolak."""
    if isinstance(waktu, datetime.datetime):
        return waktu.date()
    if isinstance(waktu, datetime.date):
        return waktu
    teks = _teks(waktu)
    for fmt in ("%Y.%m.%d %H:%M:%S.%f", "%Y.%m.%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(teks, fmt).date()
        except ValueError:
            pass
    lokasi = f" in '{path.name}', row {nomor}" if path is not None else ""
    sys.exit(f"STOP: invalid {kolom} value{lokasi}: {teks!r}")


def _kunci_tanggal(v):
    """Supaya sorted() tidak pernah meledak kalau ada tanggal yang gagal diparse
    (date dan str tidak bisa dibandingkan langsung)."""
    return v if isinstance(v, datetime.date) else datetime.date.min


class _DiskDeals:
    """Deal IDs already kept, stored on disk so multi-GB uploads stay bounded."""

    def __init__(self, path):
        self.db = sqlite3.connect(path)
        self.db.executescript("""
            PRAGMA journal_mode=OFF;
            PRAGMA synchronous=OFF;
            PRAGMA temp_store=FILE;
            CREATE TABLE deals (deal TEXT PRIMARY KEY) WITHOUT ROWID;
        """)
        self._pending = 0

    def add_deal(self, deal):
        """True when this non-empty Deal ID is seen for the first time."""
        if not deal:
            return True
        inserted = self.db.execute(
            "INSERT OR IGNORE INTO deals(deal) VALUES(?)", (deal,)).rowcount == 1
        self._pending += 1
        if self._pending >= _DEAL_COMMIT_EVERY:
            self.commit()
        return inserted

    def commit(self):
        self.db.commit()
        self._pending = 0

    def close(self):
        self.db.close()


def _pemisah_mt4(baris: str):
    if baris.count(";") > 0 and baris.count(";") >= baris.count("\t"):
        return ";"
    if baris.count("\t") > 0:
        return "\t"
    return None


def _dua_baris_awal(path: Path):
    """Dua baris non-kosong pertama. ('csv', [baris, baris]) atau ('xlsx', [tuple, tuple])."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        wb = None
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.worksheets[0] if wb.worksheets else None
            rows = []
            if ws is not None:
                for row in ws.iter_rows(max_row=2, values_only=True):
                    rows.append(tuple(_teks(v) for v in row))
            while len(rows) < 2:
                rows.append(())
            return "xlsx", rows
        except Exception:
            return "csv", []
        finally:
            if wb is not None:
                wb.close()
    enc = _deteksi_encoding(path)
    lines = []
    with open(path, encoding=enc, newline="") as f:
        for line in f:
            if line.strip():
                lines.append(line.strip("\ufeff\r\n"))
            if len(lines) == 2:
                break
    return "csv", lines


def _info_mt4(path: Path):
    """None kalau bukan Raw Report. sys.exit kalau judulnya cocok tapi headernya rusak."""
    kind, awal = _dua_baris_awal(path)
    if kind == "xlsx":
        if not awal or not awal[0]:
            return None
        cocok = _MT4_JUDUL.match(awal[0][0].strip())
        if not cocok:
            return None
        header = list(awal[1])
        pemisah = None
    else:
        if len(awal) < 2:
            return None
        cocok = _MT4_JUDUL.match(awal[0].strip())
        if not cocok:
            return None
        pemisah = _pemisah_mt4(awal[1])
        if not pemisah:
            sys.exit(f"STOP: '{path.name}' looks like an MT4 Raw Report but the header "
                     "is not semicolon- or tab-separated.")
        header = next(csv.reader([awal[1]], delimiter=pemisah))
    header = [_teks(h) for h in header]
    hilang = [k for k in _MT4_WAJIB if k not in header]
    if hilang:
        sys.exit(f"STOP: '{path.name}' looks like an MT4 Raw Report but is missing "
                 f"columns: {hilang}")
    mulai = datetime.datetime.strptime(cocok.group(1), "%Y.%m.%d").date()
    akhir = datetime.datetime.strptime(cocok.group(2), "%Y.%m.%d").date()
    if akhir < mulai:
        sys.exit(f"STOP: MT4 Raw Report period is backwards in '{path.name}': "
                 f"{mulai} to {akhir}")
    return {"awal": mulai, "akhir": akhir, "pemisah": pemisah,
            "header": header, "kind": kind}


def _macam_mt4(tipe: str, comment: str) -> str:
    """trade / balance / cancelled / other.

    Balance dan credit (CRM swap, deposit, withdrawal, transfer) ikut aturan
    MT5: Entry selain in/out tidak masuk sheet transaksi. KVB mengkonfirmasi
    aturan ini tetap dipakai.
    """
    t = tipe.strip().lower()
    if t in ("balance", "credit"):
        return "balance"
    if t in ("buy", "sell"):
        return "trade"
    if "limit" in t or "stop" in t or comment.strip().lower() == "cancelled":
        return "cancelled"
    return "other"


def _angka_longgar(v) -> float:
    """Angka pada baris yang dibuang (balance/footer). Kosong atau rusak = 0."""
    teks = _teks(v)
    if not teks:
        return 0.0
    try:
        angka = float(teks)
    except ValueError:
        return 0.0
    return angka if math.isfinite(angka) else 0.0


def _iter_mt4(path: Path, info: dict):
    """(nomor_baris, row) untuk isi file, tanpa judul dan header."""
    if info["kind"] == "xlsx":
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            it = wb.worksheets[0].iter_rows(values_only=True)
            next(it, None)
            next(it, None)
            for nomor, row in enumerate(it, start=3):
                if any(v is not None and _teks(v) != "" for v in row):
                    yield nomor, list(row)
        finally:
            wb.close()
        return
    enc = _deteksi_encoding(path)
    with open(path, encoding=enc, newline="") as f:
        reader = csv.reader(f, delimiter=info["pemisah"])
        next(reader, None)
        next(reader, None)
        nomor = 2
        for row in reader:
            nomor += 1
            if any(_teks(c) for c in row):
                yield nomor, row


def _mata_uang_mt4(path: Path, mt4_currency):
    """USD kalau nama file tidak ada di peta. Nilai lain selain USD/USC ditolak.

    Entri untuk file yang bukan Raw Report diabaikan oleh pemanggil: fungsi ini
    hanya dipanggil untuk file MT4.
    """
    table = mt4_currency or {}
    if path.name not in table:
        return "USD", False
    mentah = table[path.name]
    chosen = str(mentah).strip().upper()
    if chosen not in _MT4_MATA_UANG:
        sys.exit(f"STOP: MT4 currency for '{path.name}' must be USD or USC, not {mentah!r}.")
    return chosen, True


def _normalisasi_mt4(path: Path, info: dict, deals, pakai_baris, currency):
    """Satu Raw Report -> baris out saja, pada Close Time.

    KVB menganggap close kemarin dan open hari ini sama, jadi kaki open tidak
    ditulis. Uang tiket (Profit, Swap, Commission, Taxes sebagai Fee) ada di
    baris out. Agent dibaca hanya untuk catatan Verifikasi dan tidak ikut
    dijumlah ke Commission, Fee, atau Profit. Deal ID dideduplikasi di database
    terpisah dari MT5.
    """
    header = info["header"]
    idx = {name: i for i, name in enumerate(header)}
    total = n_out = n_bal = n_can = n_foot = n_lain = n_dup = 0
    profit_balance = agent_total = 0.0

    def sel(row, nama):
        i = idx.get(nama)
        if i is None or i >= len(row):
            return ""
        return row[i]

    for nomor, row in _iter_mt4(path, info):
        total += 1
        if len(row) < len(header):
            row = list(row) + [""] * (len(header) - len(row))
        deal_id = _teks(sel(row, "Deal"))
        login = _teks(sel(row, "Login"))
        if not deal_id and not login:
            n_foot += 1
            continue
        tipe = _teks(sel(row, "Type"))
        comment = _teks(sel(row, "Comment"))
        macam = _macam_mt4(tipe, comment)
        if macam == "balance":
            n_bal += 1
            profit_balance += _angka_longgar(sel(row, "Profit"))
            continue
        if macam == "cancelled":
            n_can += 1
            continue
        if macam != "trade":
            n_lain += 1
            continue
        if deal_id and not deals.add_deal(deal_id):
            n_dup += 1
            continue
        if not login:
            sys.exit(f"STOP: blank Login in '{path.name}', row {nomor}")
        tutup = tanggal_dari_waktu(sel(row, "Close Time"), path, nomor, "Close Time")
        volume = _angka(sel(row, "Volume"), path, nomor, "Volume")
        commission = _angka(sel(row, "Commission"), path, nomor, "Commission")
        fee = _angka(sel(row, "Taxes"), path, nomor, "Taxes")
        swap = _angka(sel(row, "Swap"), path, nomor, "Swap")
        profit = _angka(sel(row, "Profit"), path, nomor, "Profit")
        agent = _angka(sel(row, "Agent"), path, nomor, "Agent")
        symbol = _teks(sel(row, "Symbol")).upper()
        n_out += 1
        agent_total += agent
        pakai_baris({
            "Entry": "out",
            "Login": login,
            "Date": tutup,
            "Type": tipe.strip().lower(),
            "Symbol": symbol,
            "Volume": volume,
            "Commission": commission,
            "Fee": fee,
            "Swap": swap,
            "Profit": profit,
            "Currency": currency,
        }, "mt4")
    deals.commit()
    return {"nama": path.name, "platform": "mt4", "total": total,
            "out": n_out, "balance": n_bal, "cancelled": n_can,
            "footer": n_foot, "lain": n_lain, "duplikat": n_dup,
            "balance_profit": profit_balance, "agent": agent_total,
            "currency": currency, "sheets": {}, "sheets_dilewati": []}


def _normalisasi_satu_file(path: Path, deals, pakai_baris):
    """Baca satu file; tiap baris Entry 'out' atau 'in' dikirim ke pakai_baris
    (dengan b['Entry'] = 'out'/'in'). Baris dengan Entry lain/kosong dibuang.
    Deal ID yang sudah pernah muncul (file manapun, jenis baris manapun) dibuang.
    Login kosong ditolak: itu kunci grouping, dan akun tak dikenal akan
    tertumpuk jadi satu baris kalau dibiarkan."""
    idx, rows, info = baca_file(path)
    total = n_out = n_in = lain = dup_out = dup_in = 0
    for nomor, row in rows:
        total += 1
        if len(row) != len(idx):
            sys.exit(f"STOP: invalid column count in '{path.name}', row {nomor}: "
                     f"expected {len(idx)} columns, found {len(row)} columns")
        entry = _teks(row[idx["Entry"]]).lower()
        if entry not in ENTRY_DIPAKAI:
            lain += 1
            continue
        if entry == "in":
            n_in += 1
        else:
            n_out += 1
        deal_id = _teks(row[idx["Deal"]])
        if deal_id and not deals.add_deal(deal_id):
            if entry == "in":
                dup_in += 1
            else:
                dup_out += 1
            continue
        login = _teks(row[idx["Login"]])
        tanggal = tanggal_dari_waktu(row[idx["Time"]], path, nomor)
        # Volume before the money columns, matching the source-file order, so a
        # bad Volume is reported before a bad Fee on the same row.
        volume = _angka(row[idx["Volume"]], path, nomor, "Volume")
        commission = _angka(row[idx["Commission"]], path, nomor, "Commission")
        fee = _angka(row[idx["Fee"]], path, nomor, "Fee")
        swap = _angka(row[idx["Swap"]], path, nomor, "Swap")
        profit = _angka(row[idx["Profit"]], path, nomor, "Profit")
        if not login and entry == "in" and any((commission, fee, swap)):
            sys.exit(f"STOP: blank Login on Entry = 'in' with financial charges "
                     f"in '{path.name}', row {nomor}")
        if not login:
            sys.exit(f"STOP: blank Login in '{path.name}', row {nomor}")
        pakai_baris({
            "Entry": entry,
            "Login": login,
            "Date": tanggal,
            "Type": _teks(row[idx["Type"]]),
            "Symbol": _teks(row[idx["Symbol"]]),
            "Volume": volume,
            "Commission": commission,
            "Fee": fee,
            "Swap": swap,
            "Profit": profit,
            "Currency": _teks(row[idx["Currency"]]),
        })
    deals.commit()
    return {"nama": path.name, "total": total, "out": n_out, "in": n_in,
            "lain": lain, "dup_out": dup_out, "dup_in": dup_in,
            "duplikat": dup_out + dup_in, "sheets": info["dibaca"],
            "sheets_dilewati": info["dilewati"]}


def _tambah_agregat(kelompok, b, level, pakai_fx=False):
    """Tambahkan satu baris ke agregat harian atau bulanan.

    Platform ikut kunci hanya kalau barisnya membawanya (ada file MT4). Upload
    MT5 saja tetap Login+periode+Type+Symbol+Currency.
    """
    tgl = b["Date"]
    periode = tgl.replace(day=1) if level == "bulan" else tgl
    platform = b.get("Platform")
    kunci = (b["Login"], periode, b["Type"], b["Symbol"], b["Currency"])
    if platform:
        kunci = (platform,) + kunci
    if kunci not in kelompok:
        kelompok[kunci] = {"Login": b["Login"], "Periode": periode,
                           "Type": b["Type"], "Symbol": b["Symbol"],
                           "Deals": 0, "Volume": 0.0, "Commission": 0.0, "Fee": 0.0,
                           "Swap": 0.0, "Profit": 0.0, "Currency": b["Currency"],
                           "_kurs_kurang": 0, "_ada_usd_valid": False}
        if platform:
            kelompok[kunci]["Platform"] = platform
        if pakai_fx:
            kelompok[kunci].update({name: 0.0 for name in KOLOM_USD})
    d = kelompok[kunci]
    d["Deals"] += 1
    for kol in KOLOM_JUMLAH:
        d[kol] += b[kol]
    if pakai_fx:
        if b.get("_kurs_kurang"):
            d["_kurs_kurang"] += 1
        else:
            d["_ada_usd_valid"] = True
            for kol in KOLOM_USD:
                d[kol] += b[kol]


def _hasil_agregat(kelompok):

    def kunci_urut(d):
        try:
            login_num = int(d["Login"])
        except ValueError:
            login_num = 0
        platform_rank = {"MT5": 0, "MT4": 1}.get(d.get("Platform"), 0)
        return (platform_rank, login_num, _kunci_tanggal(d["Periode"]),
                d["Type"], d["Symbol"])

    hasil = list(kelompok.values())
    hasil.sort(key=kunci_urut)
    return hasil


def proses(paths_in, path_out: Path, allow_mt4=False, mt4_currency=None) -> dict:
    """Process Deals History plus optional Client Equity FX into two XLSX files in ZIP.
    'out' and 'in' rows go to separate sheets of each workbook.
    MT4 Raw Report files are accepted only when allow_mt4 is true (KVB).
    mt4_currency maps an MT4 file name to USD or USC. Missing names default to USD.
    Names that are not an MT4 Raw Report are ignored."""
    paths_in = [Path(p) for p in paths_in]
    path_out = Path(path_out)
    fx_paths = [p for p in paths_in if _adalah_file_fx(p)]
    deal_paths = [p for p in paths_in if p not in fx_paths]
    if len(fx_paths) > 1:
        sys.exit("STOP: upload only one Client Equity FX workbook per run.")
    if not deal_paths:
        sys.exit("STOP: no Deals History file was uploaded; only Client Equity FX file(s) were found.")
    klasifikasi = [(path, _info_mt4(path)) for path in deal_paths]
    mt4_paths = [path for path, info in klasifikasi if info]
    if mt4_paths and not allow_mt4:
        names = ", ".join(path.name for path in mt4_paths)
        sys.exit(
            "STOP: DPM Deal Segregator accepts MT5 Deals History only. "
            f"MT4 Raw Report is not used on this page: {names}. "
            "Upload MT4 on the KVB Deal Segregator."
        )
    fx = {}
    for path in fx_paths:
        fx.update(baca_fx(path))
    pakai_fx = bool(fx_paths)
    kurs_kurang = defaultdict(int)

    # Satu agregat. Platform ada di kunci hanya kalau run ini punya file MT4,
    # supaya workbook MT5 saja tetap identik.
    gabung = bool(mt4_paths)
    harian = {"out": {}, "in": {}}
    bulanan = {"out": {}, "in": {}}
    login_unik = set()
    login_mt4 = set()
    dipakai = {"out": 0, "in": 0}
    dipakai_mt4 = {"out": 0, "in": 0}
    jumlah = {"out": defaultdict(float), "in": defaultdict(float)}
    jumlah_mt4 = {"out": defaultdict(float), "in": defaultdict(float)}

    def pakai_baris(b, platform="mt5"):
        if gabung:
            b["Platform"] = "MT4" if platform == "mt4" else "MT5"
        if platform == "mt4":
            dipakai_t, jumlah_t, login_t = dipakai_mt4, jumlah_mt4, login_mt4
        else:
            dipakai_t, jumlah_t, login_t = dipakai, jumlah, login_unik
        e = b["Entry"]
        dipakai_t[e] += 1
        login_t.add(b["Login"])
        for kol in KOLOM_JUMLAH:
            jumlah_t[e][kol] += b[kol]
        if pakai_fx:
            marker = object()
            key = (b["Date"], _norm(b["Currency"]))
            rate = fx.get(key, marker)
            if rate is marker or rate == 0:
                b["_kurs_kurang"] = True
                kurs_kurang[key] += 1
                for name in KOLOM_USD:
                    b[name] = None
            else:
                b["_kurs_kurang"] = False
                pembagi = 1.0 if rate is None else rate
                for source, target in zip(("Commission", "Fee", "Swap", "Profit"), KOLOM_USD):
                    b[target] = b[source] / pembagi
        _tambah_agregat(harian[e], b, "hari", pakai_fx)
        _tambah_agregat(bulanan[e], b, "bulan", pakai_fx)

    path_out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=path_out.parent) as state_dir:
        deals = _DiskDeals(Path(state_dir) / "deals.sqlite")
        deals_mt4 = (_DiskDeals(Path(state_dir) / "deals-mt4.sqlite")
                     if mt4_paths else None)
        try:
            per_file = []
            for path, info in klasifikasi:
                if info:
                    currency, selected = _mata_uang_mt4(path, mt4_currency)
                    rec = _normalisasi_mt4(path, info, deals_mt4, pakai_baris, currency)
                    rec["currency_selected"] = selected
                    per_file.append(rec)
                else:
                    per_file.append(_normalisasi_satu_file(path, deals, pakai_baris))
        finally:
            deals.close()
            if deals_mt4 is not None:
                deals_mt4.close()

    berkas_mt5 = [r for r in per_file if r.get("platform") != "mt4"]
    berkas_mt4 = [r for r in per_file if r.get("platform") == "mt4"]
    ada_mt5, ada_mt4 = bool(berkas_mt5), bool(berkas_mt4)
    total = sum(r["total"] for r in berkas_mt5)
    entry_lain = sum(r["lain"] for r in berkas_mt5)
    dup_out = sum(r["dup_out"] for r in berkas_mt5)
    dup_in = sum(r["dup_in"] for r in berkas_mt5)
    duplikat = dup_out + dup_in
    if not (dipakai["out"] or dipakai["in"] or dipakai_mt4["out"] or dipakai_mt4["in"]):
        if not ada_mt4:
            sys.exit(
                "STOP: no row with Entry = 'out' or 'in' was found.\n"
                f"Across {len(berkas_mt5)} Deals History file(s), {total} rows in total, "
                f"all {entry_lain} with another/empty Entry.\n"
                "Check that these files are the correct Deals History exports."
            )

    h_out, h_in = _hasil_agregat(harian["out"]), _hasil_agregat(harian["in"])
    b_out, b_in = _hasil_agregat(bulanan["out"]), _hasil_agregat(bulanan["in"])

    ringkasan = {"Uploaded files": len(paths_in)}
    yellow_labels = set()
    for r in per_file:
        if r.get("platform") == "mt4":
            how = ("selected for this file" if r["currency_selected"]
                   else "default; no currency was selected for this file")
            ringkasan[f"  - {r['nama']} (MT4 Raw Report)"] = (
                f"{r['total']} rows, {r['out']} closed buy/sell, "
                f"currency {r['currency']} ({how}), "
                f"{r['balance']} balance/credit dropped, {r['cancelled']} cancelled pending dropped, "
                f"{r['footer']} footer/summary dropped")
            continue
        ringkasan[f"  - {r['nama']}"] = (
            f"{r['total']} rows, {r['out']} 'out' + {r['in']} 'in' kept"
            + (f", {r['duplikat']} duplicates dropped" if r["duplikat"] else ""))
        for nama, n in r["sheets"].items():
            ringkasan[f"      sheet '{nama}'"] = f"{n} rows read"
        if r["sheets_dilewati"]:
            ringkasan[f"      sheets skipped in {r['nama']} (no Deals History columns)"] = \
                ", ".join(r["sheets_dilewati"])
        if not (r["out"] or r["in"]):
            ringkasan[f"  - {r['nama']}"] = f"{r['total']} rows, 0 'out' + 0 'in' (no usable rows)"
            yellow_labels.add(f"  - {r['nama']}")
    for path in fx_paths:
        ringkasan[f"  - {path.name} (Client Equity FX)"] = f"FX table, {len(fx):,} date/currency rates"
    ringkasan["USD conversion (Commission/Fee/Swap/Profit)"] = (
        f"applied using {len(fx):,} date/currency rates" if pakai_fx
        else "not applied — no Client Equity FX file uploaded")
    if kurs_kurang:
        label = "Missing FX rates (date+currency; USD cells highlighted yellow)"
        ordered = sorted(kurs_kurang, key=lambda item: (_kunci_tanggal(item[0]), item[1]))
        ringkasan[label] = "; ".join(
            f"{date} {currency} ({kurs_kurang[(date, currency)]} rows)"
            for date, currency in ordered[:15])
        yellow_labels.add(label)
    if ada_mt5 and not ada_mt4:
        ringkasan.update({
            "Total source rows (all files)": total,
            "Entry = out (before deduplication)": dipakai["out"] + dup_out,
            "Entry = in (before deduplication)": dipakai["in"] + dup_in,
            "Entry empty/other (dropped)": entry_lain,
            "Duplicate non-empty Deal IDs (dropped)": duplikat,
            "Unique 'out' rows kept": dipakai["out"],
            "Unique 'in' rows kept": dipakai["in"],
            # Label lists every grouping dimension, Currency included -- see _tambah_agregat.
            "Daily 'out' rows (Login+Date+Type+Symbol+Currency)": len(h_out),
            "Daily 'in' rows (same grouping, sheet 'Daily - In')": len(h_in),
            "Monthly Summary 'out' rows (Login+Month+Type+Symbol+Currency)": len(b_out),
            "Monthly Summary 'in' rows (same grouping, sheet 'Monthly Summary - In')": len(b_in),
            "Unique logins": len(login_unik),
            "Total Profit ('out' rows)": round(jumlah["out"]["Profit"], 2),
            "Total Profit ('in' rows)": round(jumlah["in"]["Profit"], 2),
            "Total Commission ('out' rows)": round(jumlah["out"]["Commission"], 2),
            "Total Commission ('in' rows)": round(jumlah["in"]["Commission"], 2),
            # Volume is not money; keep source precision (do not round to 2 decimals).
            "Total Volume ('out' rows)": jumlah["out"]["Volume"],
            "Total Volume ('in' rows) - same trades as 'out'; do not add the two": jumlah["in"]["Volume"],
        })
    elif ada_mt5:
        # Angka 'out'/'in' di sini tetap MT5 saja. Panjang sheet gabungan ada di blok MT4.
        ringkasan.update({
            "MT5 source rows": total,
            "Entry = out (before deduplication)": dipakai["out"] + dup_out,
            "Entry = in (before deduplication)": dipakai["in"] + dup_in,
            "Entry empty/other (dropped)": entry_lain,
            "Duplicate non-empty Deal IDs (dropped)": duplikat,
            "Unique 'out' rows kept": dipakai["out"],
            "Unique 'in' rows kept": dipakai["in"],
            "Unique logins": len(login_unik),
            "Total Profit ('out' rows)": round(jumlah["out"]["Profit"], 2),
            "Total Profit ('in' rows)": round(jumlah["in"]["Profit"], 2),
            "Total Commission ('out' rows)": round(jumlah["out"]["Commission"], 2),
            "Total Commission ('in' rows)": round(jumlah["in"]["Commission"], 2),
            "Total Volume ('out' rows)": jumlah["out"]["Volume"],
            "Total Volume ('in' rows) - same trades as 'out'; do not add the two": jumlah["in"]["Volume"],
        })
    if ada_mt4:
        ringkasan.update({
            "MT4 source rows": sum(r["total"] for r in berkas_mt4),
            "MT4 closed buy/sell kept": dipakai_mt4["out"],
            "MT4 balance/credit rows dropped (same rule as MT5 non-trade Entry)":
                sum(r["balance"] for r in berkas_mt4),
            "MT4 cancelled pending orders dropped": sum(r["cancelled"] for r in berkas_mt4),
            "MT4 footer/summary lines dropped": sum(r["footer"] for r in berkas_mt4),
            "MT4 duplicate Deal IDs dropped": sum(r["duplikat"] for r in berkas_mt4),
            "MT4 other rows dropped": sum(r["lain"] for r in berkas_mt4),
            "MT4 balance/credit Profit excluded": round(sum(r["balance_profit"] for r in berkas_mt4), 2),
            "Daily 'out' rows (Platform+Login+Date+Type+Symbol+Currency)": len(h_out),
            "Daily 'in' rows (Platform+Login+Date+Type+Symbol+Currency, sheet 'Daily - In')": len(h_in),
            "Monthly Summary 'out' rows (Platform+Login+Month+Type+Symbol+Currency)": len(b_out),
            "Monthly Summary 'in' rows (Platform+Login+Month+Type+Symbol+Currency, sheet 'Monthly Summary - In')": len(b_in),
            "Unique MT4 logins": len(login_mt4),
            "MT4 Total Profit (closed buy/sell)": round(jumlah_mt4["out"]["Profit"], 2),
            "MT4 Total Commission (closed buy/sell)": round(jumlah_mt4["out"]["Commission"], 2),
            "MT4 Total Fee (Taxes, closed buy/sell)": round(jumlah_mt4["out"]["Fee"], 2),
            "MT4 Total Swap (closed buy/sell)": round(jumlah_mt4["out"]["Swap"], 2),
            "MT4 Total Volume (closed buy/sell)": jumlah_mt4["out"]["Volume"],
            "MT4 Agent column (sum of kept rows; not included in Profit, Commission, or Fee)":
                round(sum(r["agent"] for r in berkas_mt4), 2),
            "Combined Total Profit (MT5 + MT4 out rows)":
                round(jumlah["out"]["Profit"] + jumlah_mt4["out"]["Profit"], 2),
            "Combined Total Commission (MT5 + MT4 out rows)":
                round(jumlah["out"]["Commission"] + jumlah_mt4["out"]["Commission"], 2),
            "Combined Total Fee (MT5 + MT4 out rows)":
                round(jumlah["out"]["Fee"] + jumlah_mt4["out"]["Fee"], 2),
            "Combined Total Swap (MT5 + MT4 out rows)":
                round(jumlah["out"]["Swap"] + jumlah_mt4["out"]["Swap"], 2),
            "Combined Total Volume (MT5 + MT4 out rows)":
                jumlah["out"]["Volume"] + jumlah_mt4["out"]["Volume"],
            "MT4 currency is chosen per file":
                "The Raw Report has no Currency column. Symbol suffix, the report title, "
                "login number and profit size cannot separate USD from USC, so each file "
                "is one currency selected on the KVB upload page (default USD). "
                "USC uses the same Client Equity FX conversion as an MT5 USC row. "
                "Balance and credit rows stay excluded, the same rule as MT5.",
        })

    if path_out.suffix.lower() == ".zip":
        _tulis_zip(h_out, b_out, h_in, b_in, path_out, ringkasan, pakai_fx, yellow_labels,
                   pakai_platform=ada_mt4)
    elif pakai_fx:
        sys.exit("STOP: Client Equity FX output must use a .zip filename.")
    else:
        _tulis_workbook(h_out, b_out, h_in, b_in, path_out, ringkasan,
                        pakai_platform=ada_mt4)

    result = {"file_masuk": len(paths_in), "total": total, "dipakai": dipakai["out"],
              "dipakai_in": dipakai["in"],
              "harian": len(h_out), "bulanan": len(b_out),
              "harian_in": len(h_in), "bulanan_in": len(b_in),
              "login_unik": len(login_unik)}
    if pakai_fx:
        result.update(pakai_fx=True, kurs_kurang=sum(kurs_kurang.values()))
    if ada_mt4:
        result.update(mt4_dipakai=dipakai_mt4["out"])
    return result


# ------------------------------------------------------------------------ tulis
KUNING = PatternFill("solid", fgColor="FFFF00")


def _tulis_sheet(ws, hasil, kolom_periode, fmt_periode, pakai_fx=False,
                 kolom_hasil=None, kolom_usd=None):
    """Stream rows into a write-only sheet so memory does not scale with row count."""
    if kolom_hasil is None:
        kolom_hasil = KOLOM_HASIL
    if kolom_usd is None:
        kolom_usd = KOLOM_USD
    kolom = [kolom_periode if k == "{PERIODE}" else k
             for k in kolom_hasil + (kolom_usd if pakai_fx else [])]
    usd_indexes = {kolom.index(k) + 1 for k in kolom_usd} if pakai_fx else set()
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for j, name in enumerate(kolom, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(10, len(name) + 2) + 4
    ws.freeze_panes = "A2"
    header = []
    for j, name in enumerate(kolom, start=1):
        c = WriteOnlyCell(ws, value=name)
        c.fill = header_fill
        c.font = header_font
        if j in usd_indexes:
            c.comment = Comment("Yellow means this row has no matching date/currency rate in the Client Equity FX file.",
                                "Deal Segregator")
        header.append(c)
    ws.append(header)

    for d in hasil:
        linha = []
        for j, name in enumerate(kolom, start=1):
            key = "Periode" if name == kolom_periode else name
            nilai = d[key]
            missing_fx = j in usd_indexes and d.get("_kurs_kurang")
            if missing_fx:
                nilai = None
            teks_tidak_aman = False
            if key in KOLOM_UANG or key in kolom_usd:
                if isinstance(nilai, float):
                    nilai = round(nilai, 2)
            elif key in KOLOM_TEKS and isinstance(nilai, str) and nilai.startswith(("=", "+", "-", "@")):
                # quotePrefix marks the cell literal in Excel. data_type "s" is
                # required as well: quotePrefix alone still serializes '=...' as <f>.
                teks_tidak_aman = True
            sel = WriteOnlyCell(ws, value=nilai)
            if teks_tidak_aman:
                sel.data_type = "s"
                sel.quotePrefix = True
            if missing_fx:
                sel.fill = KUNING
            if name == kolom_periode and isinstance(d["Periode"], datetime.date):
                sel.number_format = fmt_periode
            linha.append(sel)
        ws.append(linha)


def _tulis_verifikasi(ws, ringkasan: dict, yellow_labels=frozenset()):
    ws.column_dimensions["A"].width = 78
    ws.column_dimensions["B"].width = 22
    title = WriteOnlyCell(ws, value="Summary - Deal Segregator")
    title.font = Font(bold=True, size=13)
    ws.append([title])
    ws.append([])
    for label, val in ringkasan.items():
        first, second = WriteOnlyCell(ws, value=label), WriteOnlyCell(ws, value=val)
        if label in yellow_labels:
            first.fill = second.fill = KUNING
        ws.append([first, second])


def _tulis_workbook(harian, bulanan, harian_in, bulanan_in, path_out: Path, ringkasan: dict,
                    pakai_platform=False):
    """Legacy one-workbook writer kept for direct callers."""
    kolom = KOLOM_HASIL_PLATFORM if pakai_platform else None
    wb = Workbook(write_only=True)
    daily = wb.create_sheet("Daily")
    _tulis_sheet(daily, harian, "Date", "dd mmm yyyy", kolom_hasil=kolom)
    _tulis_sheet(wb.create_sheet("Daily - In"), harian_in, "Date", "dd mmm yyyy", kolom_hasil=kolom)
    monthly = wb.create_sheet("Monthly Summary")
    _tulis_sheet(monthly, bulanan, "Month", "mmm yyyy", kolom_hasil=kolom)
    _tulis_sheet(wb.create_sheet("Monthly Summary - In"), bulanan_in, "Month", "mmm yyyy",
                 kolom_hasil=kolom)
    _tulis_verifikasi(wb.create_sheet("Verifikasi"), ringkasan)
    path_out.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path_out)
    finally:
        wb.close()


def _write_split_workbook(path, title, rows, rows_in, period_name, period_format, summary,
                          pakai_fx, yellow_labels, pakai_platform=False):
    kolom = KOLOM_HASIL_PLATFORM if pakai_platform else None
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title)
    _tulis_sheet(ws, rows, period_name, period_format, pakai_fx, kolom)
    _tulis_sheet(wb.create_sheet(f"{title} - In"), rows_in, period_name, period_format,
                 pakai_fx, kolom)
    _tulis_verifikasi(wb.create_sheet("Verifikasi"), summary, yellow_labels)
    try:
        wb.save(path)
    finally:
        wb.close()


def _tulis_zip(harian, bulanan, harian_in, bulanan_in, path_out, summary, pakai_fx, yellow_labels,
               pakai_platform=False):
    path_out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=path_out.parent) as folder:
        daily = Path(folder) / "Deals - Daily.xlsx"
        monthly = Path(folder) / "Deals - Monthly Summary.xlsx"
        _write_split_workbook(daily, "Daily", harian, harian_in, "Date", "dd mmm yyyy",
                              summary, pakai_fx, yellow_labels, pakai_platform)
        _write_split_workbook(monthly, "Monthly Summary", bulanan, bulanan_in, "Month", "mmm yyyy",
                              summary, pakai_fx, yellow_labels, pakai_platform)
        with zipfile.ZipFile(path_out, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.write(daily, daily.name)
            archive.write(monthly, monthly.name)


def _mt4_currency_dari_argumen(items):
    """NAME=USD|USC, diulang. Nama yang sama: nilai terakhir yang menang."""
    hasil = {}
    for item in items or []:
        if "=" not in item:
            sys.exit(f"STOP: --mt4-currency must look like NAME=USD, not {item!r}.")
        name, cur = item.split("=", 1)
        name = name.strip()
        cur = cur.strip().upper()
        if not name or cur not in _MT4_MATA_UANG:
            sys.exit(f"STOP: --mt4-currency must look like NAME=USD or NAME=USC, not {item!r}.")
        hasil[name] = cur
    return hasil


def main():
    ap = argparse.ArgumentParser(description="Deal Segregator -- segregate Deals History per Login")
    ap.add_argument("input", nargs="+", help="one or more Deals History files (.csv, .xlsx or .xlsm)")
    ap.add_argument("-o", "--output", help="result .zip file (default: <input>-hasil.zip)")
    ap.add_argument("--allow-mt4", action="store_true",
                    help="accept an MT4 Manager Raw Report as well (KVB). DPM must leave this off.")
    ap.add_argument("--mt4-currency", action="append", default=[], metavar="NAME=USD|USC",
                    help="currency for one MT4 Raw Report, matched on the file name. "
                         "Default USD. Ignored for files that are not an MT4 Raw Report.")
    args = ap.parse_args()

    paths_in = [Path(p) for p in args.input]
    for p in paths_in:
        if not p.is_file():
            sys.exit(f"STOP: file not found: {p}")
    path_out = Path(args.output) if args.output else paths_in[0].with_name(f"{paths_in[0].stem}-hasil.zip")

    ringkasan = proses(paths_in, path_out, allow_mt4=args.allow_mt4,
                       mt4_currency=_mt4_currency_dari_argumen(args.mt4_currency))
    print("=== STAGE 1 DONE ===")
    for k, v in ringkasan.items():
        print(f"  {LABEL_RINGKASAN.get(k, k)}: {v}")
    print(f"Output: {path_out.resolve()}")


if __name__ == "__main__":
    main()
