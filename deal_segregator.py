#!/usr/bin/env python3
"""Deal Segregator -- segregasi export "Deals History" (MT5) per Login.

Jalankan:
    python3 deal_segregator.py "Juni.csv" "Juli.xlsx" -o hasil.xlsx

Bisa menerima SATU ATAU LEBIH file sumber sekaligus (.csv atau .xlsx/.xlsm) --
mis. upload Juni dan Juli sebagai dua file terpisah, hasilnya DIGABUNG jadi
satu laporan. Cocok untuk laporan yang mencakup beberapa bulan tapi
diekspor per bulan dari MT5.

Tahap 1 (dikonfirmasi user 11 Sep 2026, direvisi 12 & 12 Sep 2026):
  - Hanya baris dengan kolom "Entry" (kolom K di Excel) = "out" yang dipakai.
    Baris "out" itu yang MENUTUP posisi -> yang punya Profit realized. Baris
    "in" (buka posisi) dan baris kosong dibuang.
  - "Desk" diturunkan dari kolom "Group", prefix "real\\" dibuang saja
    (mis. "real\\DPMKT-15" -> "DPMKT-15").
  - "Country" dibiarkan APA ADANYA dari file sumber -- di file yang sudah
    diuji (31 Jul 2026) kolom ini KOSONG untuk 419.102 dari 419.133 baris.
    Keputusan user 11 Sep 2026: skip pemetaan Country dulu untuk tahap ini.
  - "Time" dipangkas jadi TANGGAL saja (jam/menit/detik dibuang).
  - Kalau ada beberapa file, baris dengan "Deal" ID yang SAMA (dari file
    manapun) hanya dihitung SEKALI -- jaga-jaga kalau dua file yang diupload
    ternyata tumpang tindih tanggalnya.
  - Hasilnya berupa SATU workbook .xlsx dengan sheet Daily, Monthly Summary,
    dan Verifikasi.
    Baris dengan kunci yang sama pada level masing-masing digabung, kolom
    numerik dijumlah, kolom "Deals" mencatat berapa baris asli yang tergabung.

File sumber MT5 "Deals History" biasanya diekspor sebagai UTF-16LE
tab-delimited (bukan CSV koma biasa) -- encoding & pemisah kolom dideteksi
otomatis dari file, jadi export UTF-8/koma biasa dan file .xlsx juga tetap terbaca.
"""
import argparse
import csv
import datetime
import math
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

KOLOM_WAJIB = ["Deal", "Login", "Group", "Country", "Time", "Type", "Entry",
               "Symbol", "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]

KOLOM_HASIL = ["Login", "Country", "Desk", "{PERIODE}", "Type", "Symbol", "Deals",
               "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]

KOLOM_JUMLAH = ["Volume", "Commission", "Fee", "Swap", "Profit"]
KOLOM_TEKS = {"Login", "Country", "Desk", "Type", "Symbol", "Currency"}

# Label Inggris untuk ringkasan yang DITAMPILKAN ke user (log job di web app).
# Kunci dict hasil tetap dipakai internal oleh pemanggil lain.
LABEL_RINGKASAN = {
    "file_masuk": "files read",
    "total": "source rows",
    "dipakai": "kept",
    "harian": "Daily rows",
    "bulanan": "Monthly Summary rows",
    "login_unik": "unique logins",
}

ENTRY_DIPAKAI = "out"


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
    wb = None
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        header = [_teks(v) for v in next(it)]
    except Exception as exc:
        if wb is not None:
            wb.close()
        detail = "the file has no header row" if isinstance(exc, StopIteration) else str(exc)
        sys.exit(f"STOP: could not read the header in '{path.name}': {detail}")
    def rows():
        try:
            for nomor, row in enumerate(it, start=2):
                if any(v is not None and _teks(v) != "" for v in row):
                    yield nomor, list(row)
        finally:
            wb.close()
    return header, rows(), wb.close


def baca_file(path: Path):
    """Baca satu file export Deals History (.csv atau .xlsx/.xlsm). -> (idx, rows)."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        header, rows, close = _baca_xlsx(path)
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
    return idx, rows


# ------------------------------------------------------------------- normalisasi
def _teks(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, str):
        return v.strip()
    return str(v).strip()


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


def tanggal_dari_waktu(waktu, path=None, nomor=None):
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
    sys.exit(f"STOP: invalid Time value{lokasi}: {teks!r}")


def _kunci_tanggal(v):
    """Supaya sorted() tidak pernah meledak kalau ada tanggal yang gagal diparse
    (date dan str tidak bisa dibandingkan langsung)."""
    return v if isinstance(v, datetime.date) else datetime.date.min


# --------------------------------------------------------------------- pipeline
def _normalisasi_satu_file(path: Path, deal_id_terpakai: set, pakai_baris):
    idx, rows = baca_file(path)
    total = entry_out = entry_in = entry_lain = duplikat = 0
    for nomor, row in rows:
        total += 1
        if len(row) != len(idx):
            sys.exit(f"STOP: invalid column count in '{path.name}', row {nomor}: "
                     f"expected {len(idx)} columns, found {len(row)} columns")
        entry = _teks(row[idx["Entry"]]).lower()
        if entry == "in":
            entry_in += 1
            continue
        if entry != ENTRY_DIPAKAI:
            entry_lain += 1
            continue
        entry_out += 1
        deal_id = _teks(row[idx["Deal"]])
        if deal_id and deal_id in deal_id_terpakai:
            duplikat += 1
            continue
        if deal_id:
            deal_id_terpakai.add(deal_id)
        pakai_baris({
            "Login": _teks(row[idx["Login"]]),
            # TODO: kalau tim Malaysia mengirim tabel referensi Login -> Country,
            # join di sini (mis. country = peta.get(login, _teks(row[idx["Country"]]))).
            "Country": _teks(row[idx["Country"]]),
            "Desk": desk_dari_group(_teks(row[idx["Group"]])),
            "Date": tanggal_dari_waktu(row[idx["Time"]], path, nomor),
            "Type": _teks(row[idx["Type"]]),
            "Symbol": _teks(row[idx["Symbol"]]),
            "Volume": _angka(row[idx["Volume"]], path, nomor, "Volume"),
            "Commission": _angka(row[idx["Commission"]], path, nomor, "Commission"),
            "Fee": _angka(row[idx["Fee"]], path, nomor, "Fee"),
            "Swap": _angka(row[idx["Swap"]], path, nomor, "Swap"),
            "Profit": _angka(row[idx["Profit"]], path, nomor, "Profit"),
            "Currency": _teks(row[idx["Currency"]]),
        })
    return {"nama": path.name, "total": total, "out": entry_out, "in": entry_in,
            "lain": entry_lain, "duplikat": duplikat}


def _tambah_agregat(kelompok, b, level):
    """Tambahkan satu baris ke agregat harian atau bulanan."""
    tgl = b["Date"]
    periode = tgl.replace(day=1) if level == "bulan" else tgl
    kunci = (b["Login"], b["Country"], b["Desk"], periode, b["Type"], b["Symbol"], b["Currency"])
    if kunci not in kelompok:
        kelompok[kunci] = {"Login": b["Login"], "Country": b["Country"], "Desk": b["Desk"],
                           "Periode": periode, "Type": b["Type"], "Symbol": b["Symbol"],
                           "Deals": 0, "Volume": 0.0, "Commission": 0.0, "Fee": 0.0,
                           "Swap": 0.0, "Profit": 0.0, "Currency": b["Currency"]}
    d = kelompok[kunci]
    d["Deals"] += 1
    for kol in KOLOM_JUMLAH:
        d[kol] += b[kol]


def _hasil_agregat(kelompok):

    def kunci_urut(d):
        try:
            login_num = int(d["Login"])
        except ValueError:
            login_num = 0
        return (login_num, d["Country"], d["Desk"], _kunci_tanggal(d["Periode"]),
                d["Type"], d["Symbol"])

    hasil = list(kelompok.values())
    hasil.sort(key=kunci_urut)
    return hasil


def proses(paths_in, path_out: Path) -> dict:
    """Jalankan Tahap 1 untuk file sumber dan tulis satu workbook .xlsx."""
    deal_id_terpakai = set()
    kelompok_harian, kelompok_bulanan = {}, {}
    login_unik, desk_unik = set(), set()
    dipakai = country_terisi = country_kosong = 0
    jumlah_profit = 0.0

    def pakai_baris(b):
        nonlocal dipakai, country_terisi, country_kosong, jumlah_profit
        dipakai += 1
        login_unik.add(b["Login"])
        desk_unik.add(b["Desk"])
        country_terisi += bool(b["Country"])
        country_kosong += not b["Country"]
        jumlah_profit += b["Profit"]
        _tambah_agregat(kelompok_harian, b, "hari")
        _tambah_agregat(kelompok_bulanan, b, "bulan")

    per_file = [_normalisasi_satu_file(Path(p), deal_id_terpakai, pakai_baris) for p in paths_in]

    total = sum(r["total"] for r in per_file)
    entry_in = sum(r["in"] for r in per_file)
    entry_lain = sum(r["lain"] for r in per_file)
    duplikat = sum(r["duplikat"] for r in per_file)
    if not dipakai:
        sys.exit(
            "STOP: no row with Entry = 'out' was found.\n"
            f"Across {len(paths_in)} file(s), {total} rows in total: {entry_in} 'in', "
            f"{entry_lain} other/empty.\n"
            "Check that these files are the correct Deals History exports."
        )

    harian = _hasil_agregat(kelompok_harian)
    bulanan = _hasil_agregat(kelompok_bulanan)

    ringkasan = {"Uploaded files": len(paths_in)}
    for r in per_file:
        ringkasan[f"  - {r['nama']}"] = (f"{r['total']} rows, {r['out']} 'out' kept"
                                          + (f", {r['duplikat']} duplicates dropped" if r["duplikat"] else ""))
    ringkasan.update({
        "Total source rows (all files)": total,
        "Entry = out (before deduplication)": dipakai + duplikat,
        "Entry = in (dropped)": entry_in,
        "Entry empty/other (dropped)": entry_lain,
        "Duplicate non-empty Deal IDs (dropped)": duplikat,
        "Unique rows kept": dipakai,
        "Daily output rows (Login+Date+Type+Symbol)": len(harian),
        "Monthly Summary output rows (Login+Month+Type+Symbol)": len(bulanan),
        "Unique logins": len(login_unik),
        "Rows with Country": country_terisi,
        "Rows without Country": country_kosong,
        "Unique desks": len(desk_unik),
        "Total Profit (all kept rows)": round(jumlah_profit, 2),
    })

    _tulis_workbook(harian, bulanan, path_out, ringkasan)

    return {"file_masuk": len(paths_in), "total": total, "dipakai": dipakai,
            "harian": len(harian), "bulanan": len(bulanan),
            "login_unik": len(login_unik)}


# ------------------------------------------------------------------------ tulis
def _tulis_sheet(ws, hasil, kolom_periode, fmt_periode):
    kolom = [kolom_periode if k == "{PERIODE}" else k for k in KOLOM_HASIL]
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for j, name in enumerate(kolom, start=1):
        c = ws.cell(row=1, column=j, value=name)
        c.fill = header_fill
        c.font = header_font

    col_periode = kolom.index(kolom_periode) + 1
    for i, d in enumerate(hasil, start=2):
        for j, name in enumerate(kolom, start=1):
            key = "Periode" if name == kolom_periode else name
            nilai = d[key]
            teks_tidak_aman = False
            if key in KOLOM_JUMLAH and isinstance(nilai, float):
                nilai = round(nilai, 2)
            elif key in KOLOM_TEKS and nilai.startswith(("=", "+", "-", "@")):
                # Excel would evaluate this as a formula; quotePrefix keeps it literal
                # without polluting the value with a visible apostrophe.
                teks_tidak_aman = True
            sel = ws.cell(row=i, column=j, value=nilai)
            if teks_tidak_aman:
                sel.quotePrefix = True
        if isinstance(d["Periode"], datetime.date):
            ws.cell(row=i, column=col_periode).number_format = fmt_periode

    ws.freeze_panes = "A2"
    for j, name in enumerate(kolom, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(10, len(name) + 2) + 4


def _tulis_verifikasi(ws, ringkasan: dict):
    ws.cell(row=1, column=1, value="Summary - Deal Segregator").font = Font(bold=True, size=13)
    for i, (label, val) in enumerate(ringkasan.items(), start=3):
        ws.cell(row=i, column=1, value=label)
        ws.cell(row=i, column=2, value=val)
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 22


def _tulis_workbook(harian, bulanan, path_out: Path, ringkasan: dict):
    wb = Workbook()
    daily = wb.active
    daily.title = "Daily"
    _tulis_sheet(daily, harian, "Date", "dd mmm yyyy")
    monthly = wb.create_sheet("Monthly Summary")
    _tulis_sheet(monthly, bulanan, "Month", "mmm yyyy")
    _tulis_verifikasi(wb.create_sheet("Verifikasi"), ringkasan)
    path_out.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path_out)
    finally:
        wb.close()


def main():
    ap = argparse.ArgumentParser(description="Deal Segregator -- segregate Deals History per Login")
    ap.add_argument("input", nargs="+", help="one or more Deals History files (.csv, .xlsx or .xlsm)")
    ap.add_argument("-o", "--output", help="result .xlsx file (default: <input>-hasil.xlsx)")
    args = ap.parse_args()

    paths_in = [Path(p) for p in args.input]
    for p in paths_in:
        if not p.is_file():
            sys.exit(f"STOP: file not found: {p}")
    path_out = Path(args.output) if args.output else paths_in[0].with_name(f"{paths_in[0].stem}-hasil.xlsx")

    ringkasan = proses(paths_in, path_out)
    print("=== STAGE 1 DONE ===")
    for k, v in ringkasan.items():
        print(f"  {LABEL_RINGKASAN.get(k, k)}: {v}")
    print(f"Output: {path_out.resolve()}")


if __name__ == "__main__":
    main()
