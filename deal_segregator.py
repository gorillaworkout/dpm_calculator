#!/usr/bin/env python3
"""Deal Segregator -- segregasi export "Deals History" (MT5) per Login.

Jalankan:
    python3 deal_segregator.py "Juni.csv" "Client Equity - FX.xlsx" -o hasil.zip

Bisa menerima SATU ATAU LEBIH file sumber sekaligus (.csv atau .xlsx/.xlsm) --
mis. upload Juni dan Juli sebagai dua file terpisah, hasilnya DIGABUNG jadi
satu laporan. Cocok untuk laporan yang mencakup beberapa bulan tapi
diekspor per bulan dari MT5.

Tahap 1 (dikonfirmasi user 11 Sep 2026, direvisi 12, 12 & 24 Sep 2026):
  - Baris output tetap satu per "out" (baris yang MENUTUP posisi -> yang
    punya Profit realized). Baris kosong/lainnya dibuang.
  - Baris "in" (buka posisi) TIDAK jadi baris sendiri, tapi Commission/Fee/
    Swap-nya DIGABUNG ke baris "out" yang sama Position-nya (dikonfirmasi
    user 24 Sep 2026: tim Malaysia menemukan sebagian Commission cuma
    tercatat di baris "in", jadi hilang kalau baris "in" dibuang begitu saja
    -- lihat _kumpulkan_komisi_in()). TIDAK menyertakan baris "in" sebagai
    baris terpisah karena Volume-nya HAMPIR SELALU SAMA dengan baris "out"
    pasangannya (position yang sama) -- kalau ikut disertakan, Volume
    laporan akan terhitung dua kali untuk hampir setiap trade.
  - "Group" dan "Country" diabaikan sepenuhnya (keputusan user 24 Sep 2026).
    Keduanya tidak wajib, tidak ditulis ke output, dan tidak menjadi grouping.
  - "Time" dipangkas jadi TANGGAL saja (jam/menit/detik dibuang).
  - Kalau ada beberapa file, baris dengan "Deal" ID yang SAMA (dari file
    manapun) hanya dihitung SEKALI -- jaga-jaga kalau dua file yang diupload
    ternyata tumpang tindih tanggalnya.
  - Hasilnya berupa .zip berisi Deals - Daily.xlsx dan
    Deals - Monthly Summary.xlsx; masing-masing punya sheet Verifikasi.
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
import tempfile
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

KOLOM_WAJIB = ["Deal", "Position", "Login", "Time", "Type", "Entry",
               "Symbol", "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]

KOLOM_HASIL = ["Login", "{PERIODE}", "Type", "Symbol", "Deals",
               "Volume", "Commission", "Fee", "Swap", "Profit", "Currency"]
KOLOM_USD = ["Commission (USD)", "Fee (USD)", "Swap (USD)", "Profit (USD)"]

KOLOM_JUMLAH = ["Volume", "Commission", "Fee", "Swap", "Profit"]
KOLOM_TEKS = {"Login", "Type", "Symbol", "Currency"}

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
FX_SHEET = "QUERY RESULT"


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


def _kunci_waktu(v):
    """Comparable full timestamp used to choose the earliest partial close."""
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime.combine(v, datetime.time())
    teks = _teks(v)
    for fmt in ("%Y.%m.%d %H:%M:%S.%f", "%Y.%m.%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(teks, fmt)
        except ValueError:
            pass
    return datetime.datetime.max


# --------------------------------------------------------------------- pipeline
def _kumpulkan_komisi_in(deal_paths, deal_id_terpakai: set):
    """PASS 1 (dikonfirmasi user 24 Sep 2026, opsi 'merge onto the closing
    row'): baca ULANG semua file Deals History, kali ini cuma mengumpulkan
    Commission/Fee/Swap dari baris Entry='in', dijumlah per Position.

    Dijalankan SEBELUM pass 'out' (_normalisasi_satu_file) supaya baris 'out'
    mana pun -- dari file manapun, dalam urutan upload apapun -- sudah bisa
    langsung menemukan komisi 'in' pasangannya sejak baris pertama diproses.
    Position bisa saja baru MUNCUL di file yang diupload BELAKANGAN (mis.
    "in" ada di file Januari, "out" di file Februari), jadi tidak cukup
    mengandalkan urutan baca satu file.

    Dedup pakai 'deal_id_terpakai' yang SAMA dengan pass 'out' (Deal ID unik
    per baris apa pun jenisnya, jadi berbagi satu set aman) -- supaya file
    yang sama diupload dua kali tidak melipatgandakan komisi.
    Login is part of the key because Position is not globally unique across
    client accounts.
    -> dict {(Login, Position): {"Commission","Fee","Swap"}}, counts by key"""
    in_extra = defaultdict(lambda: {"Commission": 0.0, "Fee": 0.0, "Swap": 0.0,
                                    "_fx_legs": []})
    n_per_posisi = defaultdict(int)
    out_pertama = {}
    out_deal_terpakai = set()
    for path in deal_paths:
        idx, rows = baca_file(path)
        for nomor, row in rows:
            if len(row) != len(idx):
                sys.exit(f"STOP: invalid column count in '{path.name}', row {nomor}: "
                         f"expected {len(idx)} columns, found {len(row)} columns")
            entry = _teks(row[idx["Entry"]]).lower()
            posisi = _teks(row[idx["Position"]])
            login = _teks(row[idx["Login"]])
            position_key = (login, posisi)
            if entry == "out" and posisi:
                if not login:
                    sys.exit(f"STOP: blank Login in '{path.name}', row {nomor}")
                deal_id = _teks(row[idx["Deal"]])

                # Match pass 2's first-seen Deal dedup. A later duplicate must
                # never become the target and then be discarded, or opening
                # costs would remain stranded in in_extra.
                if deal_id and (deal_id in deal_id_terpakai or deal_id in out_deal_terpakai):
                    continue
                if deal_id:
                    out_deal_terpakai.add(deal_id)
                candidate = (_kunci_waktu(row[idx["Time"]]), _teks(row[idx["Deal"]]),
                             path.name, nomor, path)
                if position_key not in out_pertama or candidate[:4] < out_pertama[position_key][:4]:
                    out_pertama[position_key] = candidate
            if entry != "in":
                continue
            deal_id = _teks(row[idx["Deal"]])
            if deal_id and deal_id in deal_id_terpakai:
                continue
            if deal_id:
                deal_id_terpakai.add(deal_id)
            commission = _angka(row[idx["Commission"]], path, nomor, "Commission")
            fee = _angka(row[idx["Fee"]], path, nomor, "Fee")
            swap = _angka(row[idx["Swap"]], path, nomor, "Swap")
            if not login and any((commission, fee, swap)):
                sys.exit(f"STOP: blank Login on Entry = 'in' with financial charges "
                         f"in '{path.name}', row {nomor}")
            if not posisi:
                if any((commission, fee, swap)):
                    sys.exit(f"STOP: blank Position on Entry = 'in' with financial charges "
                             f"in '{path.name}', row {nomor}")
                continue
            acc = in_extra[position_key]
            acc["Commission"] += commission
            acc["Fee"] += fee
            acc["Swap"] += swap
            acc["_fx_legs"].append({
                "Date": tanggal_dari_waktu(row[idx["Time"]], path, nomor),
                "Currency": _norm(row[idx["Currency"]]),
                "Commission": commission,
                "Fee": fee,
                "Swap": swap,
            })
            n_per_posisi[position_key] += 1
    target_out = {position_key: (candidate[4], candidate[3])
                  for position_key, candidate in out_pertama.items()}
    return in_extra, n_per_posisi, target_out


def _normalisasi_satu_file(path: Path, deal_id_terpakai: set, pakai_baris, in_extra,
                           posisi_dipakai: set, target_out):
    """PASS 2: baris 'out' seperti sebelumnya, TAPI Commission/Fee/Swap-nya
    sekarang ditambah dengan komisi 'in' pasangannya (dicari lewat
    'in_extra', hasil _kumpulkan_komisi_in, dikunci per Position).

    Kalau satu Position punya LEBIH DARI SATU baris 'out' (partial close --
    jarang, ~0,5% posisi di data uji), komisi 'in'-nya HANYA ditambahkan ke
    baris 'out' PERTAMA yang ditemukan (in_extra.pop -- entry-nya dibuang
    setelah dipakai), supaya tidak dobel ke setiap partial close. Posisi yang
    dipakai dicatat di 'posisi_dipakai' supaya sisa yang TIDAK PERNAH ketemu
    baris 'out'-nya bisa dilaporkan (lihat proses())."""
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
        posisi = _teks(row[idx["Position"]])
        login = _teks(row[idx["Login"]])
        if not login:
            sys.exit(f"STOP: blank Login in '{path.name}', row {nomor}")
        position_key = (login, posisi)
        tambahan = (in_extra.pop(position_key, None)
                    if posisi and target_out.get(position_key) == (path, nomor) else None)
        komisi = _angka(row[idx["Commission"]], path, nomor, "Commission")
        fee = _angka(row[idx["Fee"]], path, nomor, "Fee")
        swap = _angka(row[idx["Swap"]], path, nomor, "Swap")
        if tambahan:
            closing_currency = _norm(row[idx["Currency"]])
            opening_currencies = {
                leg["Currency"] for leg in tambahan["_fx_legs"]
                if any(leg[name] for name in ("Commission", "Fee", "Swap"))
            }
            if opening_currencies and opening_currencies != {closing_currency}:
                names = ", ".join(sorted(opening_currencies)) or "(blank)"
                sys.exit(f"STOP: opening Currency {names} does not match closing Currency "
                         f"{closing_currency or '(blank)'} in '{path.name}', row {nomor}")
            komisi += tambahan["Commission"]
            fee += tambahan["Fee"]
            swap += tambahan["Swap"]
            posisi_dipakai.add(position_key)
        pakai_baris({
            "Login": _teks(row[idx["Login"]]),
            "Date": tanggal_dari_waktu(row[idx["Time"]], path, nomor),
            "Type": _teks(row[idx["Type"]]),
            "Symbol": _teks(row[idx["Symbol"]]),
            "Volume": _angka(row[idx["Volume"]], path, nomor, "Volume"),
            "Commission": komisi,
            "Fee": fee,
            "Swap": swap,
            "Profit": _angka(row[idx["Profit"]], path, nomor, "Profit"),
            "Currency": _teks(row[idx["Currency"]]),
            "_opening_fx_legs": tambahan["_fx_legs"] if tambahan else [],
        })
    return {"nama": path.name, "total": total, "out": entry_out, "in": entry_in,
            "lain": entry_lain, "duplikat": duplikat}


def _tambah_agregat(kelompok, b, level, pakai_fx=False):
    """Tambahkan satu baris ke agregat harian atau bulanan."""
    tgl = b["Date"]
    periode = tgl.replace(day=1) if level == "bulan" else tgl
    kunci = (b["Login"], periode, b["Type"], b["Symbol"], b["Currency"])
    if kunci not in kelompok:
        kelompok[kunci] = {"Login": b["Login"], "Periode": periode,
                           "Type": b["Type"], "Symbol": b["Symbol"],
                           "Deals": 0, "Volume": 0.0, "Commission": 0.0, "Fee": 0.0,
                           "Swap": 0.0, "Profit": 0.0, "Currency": b["Currency"],
                           "_kurs_kurang": 0, "_ada_usd_valid": False}
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
        return (login_num, _kunci_tanggal(d["Periode"]),
                d["Type"], d["Symbol"])

    hasil = list(kelompok.values())
    hasil.sort(key=kunci_urut)
    return hasil


def proses(paths_in, path_out: Path) -> dict:
    """Process Deals History plus optional Client Equity FX into two XLSX files in ZIP."""
    paths_in = [Path(p) for p in paths_in]
    fx_paths = [p for p in paths_in if _adalah_file_fx(p)]
    deal_paths = [p for p in paths_in if p not in fx_paths]
    if len(fx_paths) > 1:
        sys.exit("STOP: upload only one Client Equity FX workbook per run.")
    if not deal_paths:
        sys.exit("STOP: no Deals History file was uploaded; only Client Equity FX file(s) were found.")
    fx = {}
    for path in fx_paths:
        fx.update(baca_fx(path))
    pakai_fx = bool(fx_paths)
    kurs_kurang = defaultdict(int)

    deal_id_terpakai = set()
    kelompok_harian, kelompok_bulanan = {}, {}
    login_unik = set()
    dipakai = 0
    jumlah_profit = 0.0

    def pakai_baris(b):
        nonlocal dipakai, jumlah_profit
        dipakai += 1
        login_unik.add(b["Login"])
        jumlah_profit += b["Profit"]
        if pakai_fx:
            marker = object()
            close_key = (b["Date"], _norm(b["Currency"]))
            rate = fx.get(close_key, marker)
            opening_legs = b.get("_opening_fx_legs", [])
            missing = []
            if rate is marker or rate == 0:
                missing.append(close_key)
            for leg in opening_legs:
                key = (leg["Date"], leg["Currency"])
                leg_rate = fx.get(key, marker)
                if leg_rate is marker or leg_rate == 0:
                    missing.append(key)
            if missing:
                b["_kurs_kurang"] = True
                for key in set(missing):
                    kurs_kurang[key] += 1
                for name in KOLOM_USD:
                    b[name] = None
            else:
                b["_kurs_kurang"] = False
                close_divisor = 1.0 if rate is None else rate
                for source, target in zip(("Commission", "Fee", "Swap", "Profit"), KOLOM_USD):
                    opening_native = sum(leg.get(source, 0.0) for leg in opening_legs)
                    value = (b[source] - opening_native) / close_divisor
                    for leg in opening_legs:
                        leg_rate = fx[(leg["Date"], leg["Currency"])]
                        value += leg.get(source, 0.0) / (1.0 if leg_rate is None else leg_rate)
                    b[target] = value
        _tambah_agregat(kelompok_harian, b, "hari", pakai_fx)
        _tambah_agregat(kelompok_bulanan, b, "bulan", pakai_fx)

    in_extra, n_in_per_posisi, target_out = _kumpulkan_komisi_in(deal_paths, deal_id_terpakai)
    posisi_dipakai = set()
    per_file = [_normalisasi_satu_file(path, deal_id_terpakai, pakai_baris, in_extra,
                                      posisi_dipakai, target_out)
                for path in deal_paths]

    total = sum(r["total"] for r in per_file)
    entry_in = sum(r["in"] for r in per_file)
    entry_lain = sum(r["lain"] for r in per_file)
    duplikat = sum(r["duplikat"] for r in per_file)
    if not dipakai:
        sys.exit(
            "STOP: no row with Entry = 'out' was found.\n"
            f"Across {len(deal_paths)} Deals History file(s), {total} rows in total: "
            f"{entry_in} 'in', {entry_lain} other/empty.\n"
            "Check that these files are the correct Deals History exports."
        )

    harian = _hasil_agregat(kelompok_harian)
    bulanan = _hasil_agregat(kelompok_bulanan)

    ringkasan = {"Uploaded files": len(paths_in)}
    for r in per_file:
        ringkasan[f"  - {r['nama']}"] = (f"{r['total']} rows, {r['out']} 'out' kept"
                                          + (f", {r['duplikat']} duplicates dropped" if r["duplikat"] else ""))
    for path in fx_paths:
        ringkasan[f"  - {path.name} (Client Equity FX)"] = f"FX table, {len(baca_fx(path)):,} date/currency rates"
    ringkasan["USD conversion (Commission/Fee/Swap/Profit)"] = (
        f"applied using {len(fx):,} date/currency rates" if pakai_fx
        else "not applied — no Client Equity FX file uploaded")
    yellow_labels = set()
    if kurs_kurang:
        label = "Missing FX rates (date+currency; USD cells highlighted yellow)"
        ordered = sorted(kurs_kurang, key=lambda item: (_kunci_tanggal(item[0]), item[1]))
        ringkasan[label] = "; ".join(
            f"{date} {currency} ({kurs_kurang[(date, currency)]} rows)"
            for date, currency in ordered[:15])
        yellow_labels.add(label)
    # Posisi yang sisa di in_extra SETELAH pass 'out' = baris 'in' yang komisinya
    # tidak pernah ketemu baris 'out' pasangannya di batch ini (posisi masih
    # terbuka, atau baris 'out'-nya ada di file yang tidak ikut diupload).
    # Dilaporkan (kuning), BUKAN error -- supaya tidak ada komisi yang hilang
    # tanpa jejak begitu batch berikutnya (yang memuat 'out'-nya) diupload.
    n_in_terpakai = sum(n_in_per_posisi[p] for p in posisi_dipakai)
    n_in_belum_ketemu = sum(n_in_per_posisi[p] for p in in_extra)
    ringkasan["Entry = in (merged onto matching Out by Position)"] = (
        f"{entry_in} seen, {n_in_terpakai} rows merged onto a matching Out row")
    if in_extra:
        label = "In-row commission with no matching Out in this upload (Position still open?)"
        ringkasan[label] = f"{len(in_extra):,} positions, {n_in_belum_ketemu:,} 'in' rows"
        yellow_labels.add(label)
    ringkasan.update({
        "Total source rows (all files)": total,
        "Entry = out (before deduplication)": dipakai + duplikat,
        "Entry empty/other (dropped)": entry_lain,
        "Duplicate non-empty Deal IDs (dropped)": duplikat,
        "Unique rows kept": dipakai,
        # Label lists every grouping dimension, Currency included -- see _tambah_agregat.
        "Daily output rows (Login+Date+Type+Symbol+Currency)": len(harian),
        "Monthly Summary output rows (Login+Month+Type+Symbol+Currency)": len(bulanan),
        "Unique logins": len(login_unik),
        "Total Profit (all kept rows)": round(jumlah_profit, 2),
    })

    if path_out.suffix.lower() == ".zip":
        _tulis_zip(harian, bulanan, path_out, ringkasan, pakai_fx, yellow_labels)
    elif pakai_fx:
        sys.exit("STOP: Client Equity FX output must use a .zip filename.")
    else:
        _tulis_workbook(harian, bulanan, path_out, ringkasan)

    result = {"file_masuk": len(paths_in), "total": total, "dipakai": dipakai,
              "harian": len(harian), "bulanan": len(bulanan),
              "login_unik": len(login_unik)}
    if pakai_fx:
        result.update(pakai_fx=True, kurs_kurang=sum(kurs_kurang.values()))
    return result


# ------------------------------------------------------------------------ tulis
KUNING = PatternFill("solid", fgColor="FFFF00")


def _tulis_sheet(ws, hasil, kolom_periode, fmt_periode, pakai_fx=False):
    kolom = [kolom_periode if k == "{PERIODE}" else k
             for k in KOLOM_HASIL + (KOLOM_USD if pakai_fx else [])]
    usd_indexes = {kolom.index(k) + 1 for k in KOLOM_USD} if pakai_fx else set()
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for j, name in enumerate(kolom, start=1):
        c = ws.cell(row=1, column=j, value=name)
        c.fill = header_fill
        c.font = header_font
        if j in usd_indexes:
            c.comment = Comment("Yellow means this row has no matching date/currency rate in the Client Equity FX file.",
                                "Deal Segregator")

    col_periode = kolom.index(kolom_periode) + 1
    for i, d in enumerate(hasil, start=2):
        for j, name in enumerate(kolom, start=1):
            key = "Periode" if name == kolom_periode else name
            nilai = d[key]
            if j in usd_indexes and d.get("_kurs_kurang"):
                nilai = None
            teks_tidak_aman = False
            if key in ("Commission", "Fee", "Swap", "Profit") or key in KOLOM_USD:
                if isinstance(nilai, float):
                    nilai = round(nilai, 2)
            elif key in KOLOM_TEKS and nilai.startswith(("=", "+", "-", "@")):
                # Force string storage; quotePrefix alone still serializes '=…' as <f>.
                teks_tidak_aman = True
            sel = ws.cell(row=i, column=j, value=nilai)
            if teks_tidak_aman:
                sel.data_type = "s"
            if j in usd_indexes and d.get("_kurs_kurang"):
                sel.fill = KUNING
        if isinstance(d["Periode"], datetime.date):
            ws.cell(row=i, column=col_periode).number_format = fmt_periode

    ws.freeze_panes = "A2"
    for j, name in enumerate(kolom, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(10, len(name) + 2) + 4


def _tulis_verifikasi(ws, ringkasan: dict, yellow_labels=frozenset()):
    ws.cell(row=1, column=1, value="Summary - Deal Segregator").font = Font(bold=True, size=13)
    for i, (label, val) in enumerate(ringkasan.items(), start=3):
        first = ws.cell(row=i, column=1, value=label)
        second = ws.cell(row=i, column=2, value=val)
        if label in yellow_labels:
            first.fill = second.fill = KUNING
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 22


def _tulis_workbook(harian, bulanan, path_out: Path, ringkasan: dict):
    """Legacy one-workbook writer kept for direct callers."""
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


def _write_split_workbook(path, title, rows, period_name, period_format, summary,
                          pakai_fx, yellow_labels):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    _tulis_sheet(ws, rows, period_name, period_format, pakai_fx)
    _tulis_verifikasi(wb.create_sheet("Verifikasi"), summary, yellow_labels)
    try:
        wb.save(path)
    finally:
        wb.close()


def _tulis_zip(harian, bulanan, path_out, summary, pakai_fx, yellow_labels):
    path_out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=path_out.parent) as folder:
        daily = Path(folder) / "Deals - Daily.xlsx"
        monthly = Path(folder) / "Deals - Monthly Summary.xlsx"
        _write_split_workbook(daily, "Daily", harian, "Date", "dd mmm yyyy",
                              summary, pakai_fx, yellow_labels)
        _write_split_workbook(monthly, "Monthly Summary", bulanan, "Month", "mmm yyyy",
                              summary, pakai_fx, yellow_labels)
        with zipfile.ZipFile(path_out, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.write(daily, daily.name)
            archive.write(monthly, monthly.name)


def main():
    ap = argparse.ArgumentParser(description="Deal Segregator -- segregate Deals History per Login")
    ap.add_argument("input", nargs="+", help="one or more Deals History files (.csv, .xlsx or .xlsm)")
    ap.add_argument("-o", "--output", help="result .zip file (default: <input>-hasil.zip)")
    args = ap.parse_args()

    paths_in = [Path(p) for p in args.input]
    for p in paths_in:
        if not p.is_file():
            sys.exit(f"STOP: file not found: {p}")
    path_out = Path(args.output) if args.output else paths_in[0].with_name(f"{paths_in[0].stem}-hasil.zip")

    ringkasan = proses(paths_in, path_out)
    print("=== STAGE 1 DONE ===")
    for k, v in ringkasan.items():
        print(f"  {LABEL_RINGKASAN.get(k, k)}: {v}")
    print(f"Output: {path_out.resolve()}")


if __name__ == "__main__":
    main()
