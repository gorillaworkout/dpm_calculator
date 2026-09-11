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
        f.seek(0)
        r = csv.reader(f, delimiter=pemisah)
        header = next(r)
        rows = [row for row in r if any((c or "").strip() for c in row)]
    return header, rows


def _baca_xlsx(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    header = [(_teks(v)) for v in next(it)]
    rows = [list(row) for row in it if any(v is not None and _teks(v) != "" for v in row)]
    wb.close()
    return header, rows


def baca_file(path: Path):
    """Baca satu file export Deals History (.csv atau .xlsx/.xlsm). -> (idx, rows)."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        header, rows = _baca_xlsx(path)
    else:
        header, rows = _baca_csv(path)
    idx = {name: i for i, name in enumerate(header)}
    hilang = [k for k in KOLOM_WAJIB if k not in idx]
    if hilang:
        sys.exit(
            f"BERHENTI: kolom wajib tidak ditemukan di '{path.name}': {hilang}\n"
            f"Header yang terbaca ({len(header)} kolom): {header}\n"
            "Pastikan file ini export 'Deals History' dari MT5, bukan file lain."
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


def _angka(v):
    if isinstance(v, (int, float)):
        return float(v)
    v = _teks(v)
    if v == "":
        return 0.0
    try:
        return float(v)
    except ValueError:
        return v  # biarkan apa adanya kalau memang bukan angka


def desk_dari_group(group: str) -> str:
    g = (group or "").strip()
    if g.lower().startswith("real\\"):
        g = g[len("real\\"):]
    return g


def tanggal_dari_waktu(waktu):
    """Terima string MT5 ('2026.07.31 01:10:35.454') ATAU objek datetime/date
    (kalau dibaca dari .xlsx yang selnya sudah bertipe tanggal). -> date.
    Fallback: teks asli kalau formatnya tidak dikenal (supaya tidak pernah
    crash, cuma tidak tersaring rapi)."""
    if isinstance(waktu, datetime.datetime):
        return waktu.date()
    if isinstance(waktu, datetime.date):
        return waktu
    bagian = _teks(waktu).split(" ", 1)[0]
    try:
        y, m, d = bagian.split(".")
        return datetime.date(int(y), int(m), int(d))
    except (ValueError, AttributeError):
        return bagian or "(tanpa tanggal)"


def _kunci_tanggal(v):
    """Supaya sorted() tidak pernah meledak kalau ada tanggal yang gagal diparse
    (date dan str tidak bisa dibandingkan langsung)."""
    return v if isinstance(v, datetime.date) else datetime.date.min


# --------------------------------------------------------------------- pipeline
def _normalisasi_satu_file(path: Path, deal_id_terpakai: set):
    idx, rows = baca_file(path)
    total = len(rows)
    out_rows = [row for row in rows if _teks(row[idx["Entry"]]).lower() == ENTRY_DIPAKAI]
    entry_in = sum(1 for row in rows if _teks(row[idx["Entry"]]).lower() == "in")
    entry_lain = total - len(out_rows) - entry_in

    baris, duplikat = [], 0
    for row in out_rows:
        deal_id = _teks(row[idx["Deal"]])
        if deal_id and deal_id in deal_id_terpakai:
            duplikat += 1
            continue
        if deal_id:
            deal_id_terpakai.add(deal_id)
        baris.append({
            "Login": _teks(row[idx["Login"]]),
            # TODO: kalau tim Malaysia mengirim tabel referensi Login -> Country,
            # join di sini (mis. country = peta.get(login, _teks(row[idx["Country"]]))).
            "Country": _teks(row[idx["Country"]]),
            "Desk": desk_dari_group(_teks(row[idx["Group"]])),
            "Date": tanggal_dari_waktu(row[idx["Time"]]),
            "Type": _teks(row[idx["Type"]]),
            "Symbol": _teks(row[idx["Symbol"]]),
            "Volume": _angka(row[idx["Volume"]]),
            "Commission": _angka(row[idx["Commission"]]),
            "Fee": _angka(row[idx["Fee"]]),
            "Swap": _angka(row[idx["Swap"]]),
            "Profit": _angka(row[idx["Profit"]]),
            "Currency": _teks(row[idx["Currency"]]),
        })
    return {"nama": path.name, "total": total, "out": len(out_rows), "in": entry_in,
            "lain": entry_lain, "duplikat": duplikat, "baris": baris}


def _agregasi(baris, level):
    """level: 'hari' (Date apa adanya) atau 'bulan' (dibulatkan ke tgl 1)."""
    kelompok, urutan = {}, []
    for b in baris:
        tgl = b["Date"]
        periode = tgl.replace(day=1) if (level == "bulan" and isinstance(tgl, datetime.date)) else tgl
        kunci = (b["Login"], b["Country"], b["Desk"], periode, b["Type"], b["Symbol"], b["Currency"])
        if kunci not in kelompok:
            kelompok[kunci] = {"Login": b["Login"], "Country": b["Country"], "Desk": b["Desk"],
                               "Periode": periode, "Type": b["Type"], "Symbol": b["Symbol"],
                               "Deals": 0, "Volume": 0.0, "Commission": 0.0, "Fee": 0.0,
                               "Swap": 0.0, "Profit": 0.0, "Currency": b["Currency"]}
            urutan.append(kunci)
        d = kelompok[kunci]
        d["Deals"] += 1
        for kol in KOLOM_JUMLAH:
            if isinstance(b[kol], float):
                d[kol] += b[kol]

    def kunci_urut(d):
        try:
            login_num = int(d["Login"])
        except ValueError:
            login_num = 0
        return (login_num, d["Country"], d["Desk"], _kunci_tanggal(d["Periode"]),
                d["Type"], d["Symbol"])

    hasil = [kelompok[k] for k in urutan]
    hasil.sort(key=kunci_urut)
    return hasil


def proses(paths_in, path_out: Path) -> dict:
    """Jalankan Tahap 1 untuk file sumber dan tulis satu workbook .xlsx."""
    deal_id_terpakai = set()
    per_file = [_normalisasi_satu_file(Path(p), deal_id_terpakai) for p in paths_in]

    total = sum(r["total"] for r in per_file)
    entry_in = sum(r["in"] for r in per_file)
    entry_lain = sum(r["lain"] for r in per_file)
    duplikat = sum(r["duplikat"] for r in per_file)
    semua_baris = [b for r in per_file for b in r["baris"]]
    dipakai = len(semua_baris)

    if not semua_baris:
        sys.exit(
            "BERHENTI: tidak ada satu pun baris dengan Entry = 'out' ditemukan.\n"
            f"Dari {len(paths_in)} file, {total} baris total: {entry_in} 'in', "
            f"{entry_lain} lainnya/kosong.\n"
            "Cek apakah file-file ini memang export Deals History yang benar."
        )

    harian = _agregasi(semua_baris, "hari")
    bulanan = _agregasi(semua_baris, "bulan")

    ringkasan = {"File yang diupload": len(paths_in)}
    for r in per_file:
        ringkasan[f"  - {r['nama']}"] = (f"{r['total']} baris, {r['out']} 'out' dipakai"
                                          + (f", {r['duplikat']} duplikat dibuang" if r["duplikat"] else ""))
    ringkasan.update({
        "Total baris (semua file)": total,
        "Entry = out (dipakai)": len(semua_baris) + duplikat,
        "Entry = in (dibuang)": entry_in,
        "Entry kosong/lainnya (dibuang)": entry_lain,
        "Duplikat antar-file (Deal ID sama, dibuang)": duplikat,
        "Baris unik dipakai": dipakai,
        "Baris pada Daily (Login+Date+Type+Symbol)": len(harian),
        "Baris pada Monthly Summary (Login+Month+Type+Symbol)": len(bulanan),
        "Login unik": len({b["Login"] for b in semua_baris}),
        "Country terisi": sum(1 for b in semua_baris if b["Country"]),
        "Country kosong": sum(1 for b in semua_baris if not b["Country"]),
        "Desk unik": len({b["Desk"] for b in semua_baris}),
        "Jumlah Profit (semua baris out)": round(sum(b["Profit"] for b in semua_baris
                                                      if isinstance(b["Profit"], float)), 2),
    })

    _tulis_workbook(harian, bulanan, path_out, ringkasan)

    return {"file_masuk": len(paths_in), "total": total, "dipakai": dipakai,
            "harian": len(harian), "bulanan": len(bulanan),
            "login_unik": len({b["Login"] for b in semua_baris})}


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
            if key in KOLOM_JUMLAH and isinstance(nilai, float):
                nilai = round(nilai, 2)
            ws.cell(row=i, column=j, value=nilai)
        if isinstance(d["Periode"], datetime.date):
            ws.cell(row=i, column=col_periode).number_format = fmt_periode

    ws.freeze_panes = "A2"
    for j, name in enumerate(kolom, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(10, len(name) + 2) + 4


def _tulis_verifikasi(ws, ringkasan: dict):
    ws.cell(row=1, column=1, value="Ringkasan - Deal Segregator").font = Font(bold=True, size=13)
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
    wb.save(path_out)


def main():
    ap = argparse.ArgumentParser(description="Deal Segregator -- segregasi Deals History per Login")
    ap.add_argument("input", nargs="+", help="satu atau lebih file Deals History (.csv atau .xlsx)")
    ap.add_argument("-o", "--output", help="file .xlsx hasil (default: <input>-hasil.xlsx)")
    args = ap.parse_args()

    paths_in = [Path(p) for p in args.input]
    for p in paths_in:
        if not p.is_file():
            sys.exit(f"BERHENTI: file tidak ditemukan: {p}")
    path_out = Path(args.output) if args.output else paths_in[0].with_name(f"{paths_in[0].stem}-hasil.xlsx")

    ringkasan = proses(paths_in, path_out)
    print("=== STAGE 1 SELESAI ===")
    for k, v in ringkasan.items():
        print(f"  {k}: {v}")
    print(f"Output: {path_out.resolve()}")


if __name__ == "__main__":
    main()
