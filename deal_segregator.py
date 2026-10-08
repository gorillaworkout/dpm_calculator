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

MT4 (KVB saja, flag --allow-mt4): dikenali dari header (Deal, Open Time, Close
Time, ...), bukan dari nama file. Dua bentuk: Raw Report berjudul (titik koma)
dan export yang header-nya di baris pertama (koma, boleh ada kolom mata uang
di ujung). Hanya buy/sell yang sudah ditutup (Close Time) yang masuk, ke sheet
yang sama dengan MT5, dengan kolom Platform. Tanpa file MT4, sheet dan angka
MT5 tidak berubah. Mata uang per baris: peta account_type (CentAccount = USC),
lalu kolom mata uang di file, lalu pilihan USD/USC per file. USC MT4 dibagi 100
(bukan kurs Client Equity FX). Agent tidak masuk ke total manapun.
"""
import argparse
import codecs
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


def _kunci_akun(v) -> str:
    """Login / account sebagai teks, tanpa '.0' dari sel Excel."""
    teks = _teks(v).strip()
    if re.fullmatch(r"\d+\.0+", teks):
        return teks.split(".", 1)[0]
    return teks


def _encoding_mt4_penuh(path: Path) -> str:
    """UTF-8 kalau seluruh file valid, kalau tidak cp1252.

    Header Juli 2026 lolos UTF-8; byte 0xD0 baru muncul di kolom Comment.
    Pemindaian ini hanya dipakai setelah file sudah dikenali sebagai MT4,
    jadi export MT5 multi-GB tidak ikut terpindai.
    """
    dasar = _deteksi_encoding(path)
    if dasar != "utf-8":
        return dasar
    dekoder = codecs.getincrementaldecoder("utf-8")()
    with open(path, "rb") as f:
        while True:
            potong = f.read(1 << 20)
            if not potong:
                try:
                    dekoder.decode(b"", final=True)
                except UnicodeDecodeError:
                    return "cp1252"
                return "utf-8"
            try:
                dekoder.decode(potong, final=False)
            except UnicodeDecodeError:
                return "cp1252"


def _parse_header_mt4(baris: str):
    """(pemisah, header) kalau baris ini header MT4, kalau tidak None."""
    calon = []
    for pemisah in (";", "\t", ","):
        if baris.count(pemisah) == 0:
            continue
        sel = next(csv.reader([baris], delimiter=pemisah))
        header = [_teks(h) for h in sel]
        if all(k in header for k in _MT4_WAJIB):
            calon.append((baris.count(pemisah), pemisah, header))
    if not calon:
        return None
    calon.sort(reverse=True)
    _, pemisah, header = calon[0]
    return pemisah, header


def _indeks_mata_uang_mt4(header):
    """Kolom trailing yang headernya kosong (atau bernama Currency).

    Layout Juli: 23 kolom biasa, satu kolom kosong, lalu mata uang tanpa judul.
    Satu sel kosong di ujung karena delimiter sisa tidak dihitung.
    """
    if not header:
        return None
    terakhir = header[-1].strip().lower()
    if terakhir == "currency":
        return len(header) - 1
    if terakhir == "" and len(header) >= 2 and header[-2].strip() == "":
        return len(header) - 1
    return None


_KOLOM_PETA = ("account", "account_type")


def _header_peta_dari_sel(cells):
    nama = [_teks(c).strip().lower() for c in cells]
    if all(k in nama for k in _KOLOM_PETA):
        return {k: nama.index(k) for k in _KOLOM_PETA}
    return None


def _adalah_peta_akun(path: Path) -> bool:
    """True kalau file ini peta account / account_type, bukan Deals History."""
    try:
        suffix = path.suffix.lower()
        if suffix in (".xlsx", ".xlsm"):
            wb = load_workbook(path, read_only=True, data_only=True)
            try:
                ws = wb.worksheets[0] if wb.worksheets else None
                if ws is None:
                    return False
                pertama = next(ws.iter_rows(max_row=1, values_only=True), None)
                return bool(pertama and _header_peta_dari_sel(pertama))
            finally:
                wb.close()
        with open(path, encoding=_encoding_mt4_penuh(path), newline="") as f:
            baris = f.readline()
        if not baris.strip():
            return False
        parsed = None
        for pemisah in (";", "\t", ","):
            if baris.count(pemisah) == 0:
                continue
            sel = next(csv.reader([baris], delimiter=pemisah))
            if _header_peta_dari_sel(sel):
                parsed = True
                break
        return bool(parsed)
    except Exception:
        return False


def baca_peta_akun(path: Path) -> dict:
    """account -> 'USD' atau 'USC'. CentAccount = USC. Tipe lain = USD.

    Dua tipe berbeda untuk akun yang sama menghentikan run.
    """
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            it = ws.iter_rows(values_only=True)
            header = next(it, None)
            if not header:
                sys.exit(f"STOP: '{path.name}' has no MT4 account-type header.")
            idx = _header_peta_dari_sel(header)
            if not idx:
                sys.exit(f"STOP: '{path.name}' needs columns account and account_type.")
            baris_sumber = ((n, row) for n, row in enumerate(it, start=2))
        except SystemExit:
            wb.close()
            raise
        # workbook stays open until the rows are consumed
        def tutup():
            wb.close()
    else:
        enc = _encoding_mt4_penuh(path)
        f = open(path, encoding=enc, newline="")
        try:
            sample = f.readline()
            pemisah = ","
            for calon in (";", "\t", ","):
                if sample.count(calon) == 0:
                    continue
                sel = next(csv.reader([sample], delimiter=calon))
                if _header_peta_dari_sel(sel):
                    pemisah = calon
                    header = sel
                    break
            else:
                f.close()
                sys.exit(f"STOP: '{path.name}' needs columns account and account_type.")
            idx = _header_peta_dari_sel(header)
            reader = csv.reader(f, delimiter=pemisah)

            def baris_sumber():
                for n, row in enumerate(reader, start=2):
                    yield n, row

            baris_sumber = baris_sumber()
        except SystemExit:
            f.close()
            raise

        def tutup():
            f.close()
    hasil = {}
    try:
        for nomor, row in baris_sumber:
            if row is None or idx["account"] >= len(row):
                continue
            akun = _kunci_akun(row[idx["account"]])
            if not akun:
                continue
            tipe = _teks(row[idx["account_type"]] if idx["account_type"] < len(row) else "")
            mata = "USC" if tipe.strip().lower() == "centaccount" else "USD"
            lama = hasil.get(akun)
            if lama is not None and lama != mata:
                sys.exit(f"STOP: account {akun} has conflicting account types in '{path.name}', "
                         f"row {nomor}.")
            hasil[akun] = mata
    finally:
        tutup()
    return hasil


def _selesaikan_mata_uang_mt4(login, kolom, fallback, akun):
    """(mata uang, sumber, beda). Sumber: mapping, column, atau fallback.

    Beda diisi kalau peta dan kolom sama-sama USD/USC tapi tidak sama.
    Peta menang.
    """
    mapped = (akun or {}).get(_kunci_akun(login))
    col = _teks(kolom).strip().upper()
    col_ok = col in _MT4_MATA_UANG
    if mapped:
        beda = (_kunci_akun(login), mapped, col) if col_ok and col != mapped else None
        return mapped, "mapping", beda
    if col_ok:
        return col, "column", None
    return fallback, "fallback", None


def _daftar_login(logins) -> str:
    items = sorted(logins, key=lambda s: (not s.isdigit(), int(s) if s.isdigit() else 0, s))
    if not items:
        return "(none)"
    shown = items[:40]
    text = ", ".join(shown)
    if len(items) > len(shown):
        text += f", and {len(items) - len(shown)} more"
    return text


def _teks_beda_mt4(items) -> str:
    unik = []
    lihat = set()
    for login, mapped, col in items:
        kunci = (str(login), mapped, col)
        if kunci in lihat:
            continue
        lihat.add(kunci)
        unik.append(f"login {login} mapping {mapped} column {col}")
    if not unik:
        return "(none)"
    shown = unik[:30]
    text = "; ".join(shown)
    if len(unik) > len(shown):
        text += f"; and {len(unik) - len(shown)} more"
    return text


def _baris_awal_csv(path: Path, enc: str):
    lines = []
    with open(path, encoding=enc, newline="") as f:
        for line in f:
            if line.strip():
                lines.append(line.strip("\ufeff\r\n"))
            if len(lines) == 2:
                break
    return lines


def _dua_baris_awal(path: Path):
    """Dua baris non-kosong pertama. ('csv', [baris, baris]) atau ('xlsx', [tuple, tuple]).

    Deteksi hanya butuh judul/header, yang ASCII. Baris data kedua pada layout
    header-first boleh berisi byte cp1252 (komentar), jadi UTF-8 yang gagal
    diulang sebagai cp1252. File utuh baru dipindai setelah dikenali sebagai MT4.
    """
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
    try:
        lines = _baris_awal_csv(path, enc)
    except UnicodeDecodeError:
        if enc not in ("utf-8", "utf-8-sig"):
            raise
        lines = _baris_awal_csv(path, "cp1252")
    return "csv", lines


def _bungkus_info_mt4(kind, header, pemisah, skip, mulai, akhir):
    return {"awal": mulai, "akhir": akhir, "pemisah": pemisah, "header": header,
            "kind": kind, "skip": skip, "currency_index": _indeks_mata_uang_mt4(header)}


def _info_mt4(path: Path):
    """None kalau bukan export MT4. sys.exit kalau judulnya cocok tapi headernya rusak.

    Dua bentuk: baris judul `Raw Report for ...` lalu header, atau header di baris
    pertama (koma, titik koma, atau tab).
    """
    kind, awal = _dua_baris_awal(path)
    if not awal or not awal[0]:
        return None
    if kind == "xlsx":
        header = [_teks(c) for c in awal[0]]
        if all(k in header for k in _MT4_WAJIB):
            return _bungkus_info_mt4("xlsx", header, None, 1, None, None)
        judul = _teks(awal[0][0]) if awal[0] else ""
        cocok = _MT4_JUDUL.match(judul.strip())
        if not cocok:
            return None
        header = [_teks(c) for c in awal[1]]
        hilang = [k for k in _MT4_WAJIB if k not in header]
        if hilang:
            sys.exit(f"STOP: '{path.name}' looks like an MT4 Raw Report but is missing "
                     f"columns: {hilang}")
        mulai = datetime.datetime.strptime(cocok.group(1), "%Y.%m.%d").date()
        akhir = datetime.datetime.strptime(cocok.group(2), "%Y.%m.%d").date()
        if akhir < mulai:
            sys.exit(f"STOP: MT4 Raw Report period is backwards in '{path.name}': "
                     f"{mulai} to {akhir}")
        return _bungkus_info_mt4("xlsx", header, None, 2, mulai, akhir)

    parsed = _parse_header_mt4(awal[0])
    if parsed:
        pemisah, header = parsed
        return _bungkus_info_mt4("csv", header, pemisah, 1, None, None)
    cocok = _MT4_JUDUL.match(awal[0].strip())
    if not cocok:
        return None
    if len(awal) < 2 or not _parse_header_mt4(awal[1]):
        sys.exit(f"STOP: '{path.name}' looks like an MT4 Raw Report but is missing "
                 f"columns: {[k for k in _MT4_WAJIB]}")
    pemisah, header = _parse_header_mt4(awal[1])
    mulai = datetime.datetime.strptime(cocok.group(1), "%Y.%m.%d").date()
    akhir = datetime.datetime.strptime(cocok.group(2), "%Y.%m.%d").date()
    if akhir < mulai:
        sys.exit(f"STOP: MT4 Raw Report period is backwards in '{path.name}': "
                 f"{mulai} to {akhir}")
    return _bungkus_info_mt4("csv", header, pemisah, 2, mulai, akhir)


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
    skip = info.get("skip", 2)
    if info["kind"] == "xlsx":
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            it = wb.worksheets[0].iter_rows(values_only=True)
            for _ in range(skip):
                next(it, None)
            for nomor, row in enumerate(it, start=skip + 1):
                if any(v is not None and _teks(v) != "" for v in row):
                    yield nomor, list(row)
        finally:
            wb.close()
        return
    enc = _encoding_mt4_penuh(path)
    with open(path, encoding=enc, newline="") as f:
        reader = csv.reader(f, delimiter=info["pemisah"])
        for _ in range(skip):
            next(reader, None)
        nomor = skip
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


def _normalisasi_mt4(path: Path, info: dict, deals, pakai_baris, currency, akun=None):
    """Satu export MT4 -> baris out saja, pada Close Time.

    Mata uang tiap baris: peta account type, lalu kolom di file, lalu pilihan
    per file. USC dibagi 100 pada Commission, Fee, Swap, dan Profit sebelum
    diagregasi. Volume tidak dibagi. Agent hanya untuk Verifikasi.
    """
    header = info["header"]
    idx = {name: i for i, name in enumerate(header)}
    kolom_uang = info.get("currency_index")
    total = n_out = n_bal = n_can = n_foot = n_lain = n_dup = 0
    profit_balance = agent_total = 0.0
    raw = {"Profit": 0.0, "Commission": 0.0, "Fee": 0.0, "Swap": 0.0}
    sumber_baris = {"mapping": 0, "column": 0, "fallback": 0}
    sumber_login = {"mapping": set(), "column": set(), "fallback": set()}
    beda = []
    dipakai_peta = False

    def sel(row, nama):
        i = idx.get(nama)
        if i is None or i >= len(row):
            return ""
        return row[i]

    def catat(login, row):
        nonlocal dipakai_peta
        mentah = row[kolom_uang] if kolom_uang is not None and kolom_uang < len(row) else ""
        mata, sumber, salah = _selesaikan_mata_uang_mt4(login, mentah, currency, akun)
        sumber_baris[sumber] += 1
        sumber_login[sumber].add(_kunci_akun(login))
        if salah:
            beda.append(salah)
        if sumber == "mapping":
            dipakai_peta = True
        return mata

    for nomor, row in _iter_mt4(path, info):
        total += 1
        if len(row) < len(header):
            row = list(row) + [""] * (len(header) - len(row))
        deal_id = _teks(sel(row, "Deal"))
        login = _teks(sel(row, "Login"))
        if not deal_id and not login:
            n_foot += 1
            continue
        mata = catat(login, row) if login else currency
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
        raw["Profit"] += profit
        raw["Commission"] += commission
        raw["Fee"] += fee
        raw["Swap"] += swap
        baris = {
            "Entry": "out",
            "Login": login,
            "Date": tutup,
            "Type": tipe.strip().lower(),
            "Symbol": _teks(sel(row, "Symbol")).upper(),
            "Volume": volume,
            "Commission": commission,
            "Fee": fee,
            "Swap": swap,
            "Profit": profit,
            "Currency": mata,
        }
        if mata == "USC":
            for kolom in ("Commission", "Fee", "Swap", "Profit"):
                baris[kolom] = baris[kolom] / 100.0
            baris["_mt4_usc"] = True
        n_out += 1
        agent_total += agent
        pakai_baris(baris, "mt4")
    deals.commit()
    return {"nama": path.name, "platform": "mt4", "total": total,
            "out": n_out, "balance": n_bal, "cancelled": n_can,
            "footer": n_foot, "lain": n_lain, "duplikat": n_dup,
            "balance_profit": profit_balance, "agent": agent_total,
            "currency": currency, "currency_column": kolom_uang is not None,
            "used_mapping": dipakai_peta, "sources": sumber_baris,
            "source_logins": sumber_login, "disagree": beda, "raw": raw,
            "sheets": {}, "sheets_dilewati": []}


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


def proses(paths_in, path_out: Path, allow_mt4=False, mt4_currency=None,
           mt4_accounts=None) -> dict:
    """Process Deals History plus optional Client Equity FX into two XLSX files in ZIP.
    'out' and 'in' rows go to separate sheets of each workbook.
    MT4 Raw Report files are accepted only when allow_mt4 is true (KVB).
    mt4_currency maps an MT4 file name to USD or USC. Missing names default to USD.
    Names that are not an MT4 Raw Report are ignored."""
    paths_in = [Path(p) for p in paths_in]
    path_out = Path(path_out)
    fx_paths = [p for p in paths_in if _adalah_file_fx(p)]
    rest = [p for p in paths_in if p not in fx_paths]
    if mt4_accounts and not allow_mt4:
        sys.exit("STOP: MT4 account types are only used on the KVB Deal Segregator (--allow-mt4).")
    peta_paths = []
    if mt4_accounts:
        peta = Path(mt4_accounts)
        if not peta.is_file():
            sys.exit(f"STOP: MT4 account-type file not found: {peta}")
        peta_paths.append(peta)
    if allow_mt4:
        sudah = {p.resolve() for p in peta_paths}
        for path in rest:
            if path.resolve() in sudah:
                continue
            if _adalah_peta_akun(path):
                peta_paths.append(path)
                sudah.add(path.resolve())
    if len(peta_paths) > 1:
        names = ", ".join(p.name for p in peta_paths)
        sys.exit(f"STOP: upload only one MT4 account-type file. Found: {names}.")
    akun = baca_peta_akun(peta_paths[0]) if peta_paths else {}
    peta_resolved = {p.resolve() for p in peta_paths}
    deal_paths = [p for p in rest if p.resolve() not in peta_resolved]
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
        if b.get("_mt4_usc"):
            # Sudah dibagi 100. Jangan pakai kurs USC di Client Equity FX.
            if pakai_fx:
                b["_kurs_kurang"] = False
                for source, target in zip(("Commission", "Fee", "Swap", "Profit"), KOLOM_USD):
                    b[target] = b[source]
        elif pakai_fx:
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
                    rec = _normalisasi_mt4(path, info, deals_mt4, pakai_baris, currency, akun)
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
            if r.get("currency_column") or r.get("used_mapping"):
                how = ("resolved per row: account type, then the currency column, "
                       "then the file choice")
            elif r["currency_selected"]:
                how = f"{r['currency']} (selected for this file)"
            else:
                how = f"{r['currency']} (default; no currency was selected for this file)"
            ringkasan[f"  - {r['nama']} (MT4 Raw Report)"] = (
                f"{r['total']} rows, {r['out']} closed buy/sell, "
                f"currency {how}, "
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
    if peta_paths:
        n_cent = sum(1 for mata in akun.values() if mata == "USC")
        ringkasan[f"  - {peta_paths[0].name} (MT4 account types)"] = (
            f"{len(akun)} accounts, {n_cent} CentAccount (USC), "
            f"{len(akun) - n_cent} other types (USD)")
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
            "MT4 Total Profit raw (before USC ÷100)": round(sum(r["raw"]["Profit"] for r in berkas_mt4), 2),
            "MT4 Total Commission raw (before USC ÷100)":
                round(sum(r["raw"]["Commission"] for r in berkas_mt4), 2),
            "MT4 Total Fee raw (before USC ÷100)": round(sum(r["raw"]["Fee"] for r in berkas_mt4), 2),
            "MT4 Total Swap raw (before USC ÷100)": round(sum(r["raw"]["Swap"] for r in berkas_mt4), 2),
            "MT4 currency from account-type mapping (rows)":
                sum(r["sources"]["mapping"] for r in berkas_mt4),
            "MT4 currency from account-type mapping (logins)":
                len(set().union(*(r["source_logins"]["mapping"] for r in berkas_mt4))),
            "MT4 currency from the currency column (rows)":
                sum(r["sources"]["column"] for r in berkas_mt4),
            "MT4 currency from the currency column (logins)":
                len(set().union(*(r["source_logins"]["column"] for r in berkas_mt4))),
            "MT4 currency from the per-file USD/USC choice (rows)":
                sum(r["sources"]["fallback"] for r in berkas_mt4),
            "MT4 currency from the per-file USD/USC choice (logins)":
                len(set().union(*(r["source_logins"]["fallback"] for r in berkas_mt4))),
            "MT4 logins still on the per-file currency fallback": _daftar_login(
                set().union(*(r["source_logins"]["fallback"] for r in berkas_mt4))),
            "MT4 account type and currency column disagree": _teks_beda_mt4(
                [item for r in berkas_mt4 for item in r["disagree"]]),
            "MT4 USC ÷100":
                "Commission, Fee, Swap and Profit on USC rows are divided by 100 into USD. "
                "Volume is not divided. Client Equity FX is not applied to MT4 USC.",
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
                "Each row uses the account-type mapping first (CentAccount is USC, every "
                "other type is USD), then a USD or USC value in the file, then the per-file "
                "choice (default USD). Symbol suffix, the report title, login number and "
                "profit size cannot separate USD from USC. MT4 USC amounts on the sheet are "
                "divided by 100. Balance and credit rows stay excluded, the same rule as MT5.",
        })
        label_beda = "MT4 account type and currency column disagree"
        if ringkasan.get(label_beda) not in ("(none)", "", None):
            yellow_labels.add(label_beda)

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
    ap.add_argument("--mt4-accounts",
                    help="optional account,account_type file. CentAccount is USC; "
                         "every other type is USD. Used only with --allow-mt4.")
    args = ap.parse_args()

    paths_in = [Path(p) for p in args.input]
    for p in paths_in:
        if not p.is_file():
            sys.exit(f"STOP: file not found: {p}")
    path_out = Path(args.output) if args.output else paths_in[0].with_name(f"{paths_in[0].stem}-hasil.zip")

    ringkasan = proses(paths_in, path_out, allow_mt4=args.allow_mt4,
                       mt4_currency=_mt4_currency_dari_argumen(args.mt4_currency),
                       mt4_accounts=args.mt4_accounts)
    print("=== STAGE 1 DONE ===")
    for k, v in ringkasan.items():
        print(f"  {LABEL_RINGKASAN.get(k, k)}: {v}")
    print(f"Output: {path_out.resolve()}")


if __name__ == "__main__":
    main()
