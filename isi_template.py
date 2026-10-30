#!/usr/bin/env python3
"""Isi 'Template D&W.xlsx' dengan data dari workbook sumber.

Kolom dicocokkan berdasarkan NAMA HEADER, bukan posisi -- jadi urutan kolom di
file sumber boleh berbeda, dan kolom yang tidak ada di template diabaikan.

Pakai:
    python3 isi_template.py "sumber.xlsx"
    python3 isi_template.py "sumber.xlsx" --bulan 2026-06     # hanya Juni 2026
    python3 isi_template.py "sumber.xlsx" -t "Template D&W.xlsx" -o "Juni 2026.xlsx"
    python3 isi_template.py "sumber.xlsx" --sheet-d "Deposits" --sheet-w "Withdrawals"

File sumber dan template tidak diubah -- hasilnya file baru.
"""

import argparse
import datetime
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl belum terpasang:  python3 -m pip install --user openpyxl")

# tanggal_util.py (penentuan tanggal per baris, dipakai bersama hitung_dw.py) ada di
# folder yang sama. Folder skrip belum tentu ada di sys.path kalau file ini dimuat
# lewat importlib (test, app), jadi ditambahkan dulu.
_FOLDER_SKRIP = str(Path(__file__).resolve().parent)
if _FOLDER_SKRIP not in sys.path:
    sys.path.insert(0, _FOLDER_SKRIP)
import tanggal_util as tgl  # noqa: E402

KOLOM_TGL = ("PAID DATE", "COMPLETED DATE", "SETTLEMENT DATE", "APPLY DATE")

# Fund Transfer Table: dua sisi bersebelahan dengan nama kolom yang SAMA
# (金额, 手续费, Xero (USD) muncul dua kali). Jadi kolom dicocokkan per SISI --
# batas sisi = kolom '收款日期'. Kalau dicocokkan pakai nama saja, kunci yang
# kembar saling menimpa dan sisi 汇出 ikut terisi angka sisi 收款.
FTT_SHEET = "Fund Transfer Table"
FTT_BATAS_SISI = "收款日期"
FTT_BARIS_HEADER = 3          # di template: baris 2-3 header, data mulai baris 4

# Saldo pembuka: diambil dari sheet 'J Wallet' dan 'Payment Channel Balance'
# yang ada di workbook sumber, lalu ditulis ke sheet 'Opening Balance' template.
# Tanpa ini saldo di 'Channel Balance' dan 'J Wallet (calc)' mulai dari nol.
OB_SHEET = "Opening Balance"
OB_BARIS_HEADER = 2
DIHITUNG = {"HANDLING FEE", "XERO RATE", "XERO USD", "FOREX GAIN/LOSS",
            "FOREX GAIN/-LOSS", "CURRENCY GAIN/LOSS"}


def norm(v):
    return " ".join(str(v if v is not None else "").strip().upper().split())


def cari_sheet(wb, kandidat):
    for x in wb.sheetnames:
        if norm(x) in kandidat:
            return x
    return None


# Batas akal tanggal. Di luar ini (mis. 1970-01-01 = sel tanggal 'nol' dari
# sistem sumber) tanggalnya dianggap TIDAK SAH: barisnya tidak masuk bulan mana
# pun, tapi DILAPORKAN (lihat tanggal_tidak_sah) -- bukan dibuang diam-diam.
TAHUN_MIN, TAHUN_MAX = 1990, 2100


def _tanggal_mentah(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    # Angka serial Excel (sel tanpa format tanggal, mis. 46235 = 1 Agu 2026)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            from openpyxl.utils.datetime import from_excel
            d = from_excel(v)
        except (ValueError, OverflowError, TypeError):
            return None
        if isinstance(d, datetime.datetime):
            return d.date()
    return None


def tanggal(v):
    d = _tanggal_mentah(v)
    if d is None or not TAHUN_MIN <= d.year <= TAHUN_MAX:
        return None
    return d


def tanggal_tidak_sah(v):
    """True kalau sel berisi tanggal/serial yang terbaca tapi di luar batas akal."""
    d = _tanggal_mentah(v)
    return d is not None and not TAHUN_MIN <= d.year <= TAHUN_MAX


def kunci_hdr(v):
    """Kunci header yang tahan header DUA BAHASA.

    Sebagian workbook menulis header sebagai '收款日期\nDeposit Date' (Mandarin,
    baris baru, lalu Inggris), sebagian lagi cuma '收款日期'. Kalau dicocokkan
    persis, kolomnya tidak ketemu dan seluruh sheet dilewati diam-diam --
    itu yang terjadi pada '20.08 Bayu ... _3 Tabs.xlsx'. Jadi yang dipakai
    cuma baris PERTAMA dari isi sel."""
    t = str(v if v is not None else "").strip()
    if not t:
        return ""
    return " ".join(t.splitlines()[0].strip().upper().split())


def peta_ftt(baris_header):
    """{(sisi, kunci header): indeks kolom} untuk satu baris header FTT."""
    batas = next((i for i, v in enumerate(baris_header)
                  if kunci_hdr(v) == kunci_hdr(FTT_BATAS_SISI)), None)
    peta = {}
    for i, v in enumerate(baris_header):
        if v in (None, ""):
            continue
        sisi = "in" if (batas is not None and i >= batas) else "out"
        peta.setdefault((sisi, kunci_hdr(v)), i)
    return peta


# Kolom tanggal sisi 汇出 (dana keluar) -- ini yang dipakai untuk menentukan
# sebuah pemindahan dana masuk bulan laporan yang mana. Sama dengan basis yang
# dipakai hitung_dw.baca_fund_transfer, jadi kedua tahap sepakat.
FTT_KOLOM_TGL = ("汇款日期", "DATE")

# Baris yang tanggalnya di luar TAHUN_MIN..TAHUN_MAX: (sheet, baris, nilai, info)
TANGGAL_TIDAK_SAH = []
SHEET_TGL_TIDAK_SAH = "Verifikasi Tanggal"

# Satu tanggal_util.Ringkasan per sheet (D, W) yang diproses -- diisi main().
RINGKASAN = []


def laporkan_ringkasan(ring, utama, saring, bulan_arg):
    """Cetak ke log (= log job web) apa yang dilewati / diselamatkan di satu sheet.
    Tidak ada baris yang boleh hilang tanpa jejak: semuanya juga masuk sheet
    'Date Check' di file hasil."""
    if ring.total == 0:
        return
    t = ring.sheet
    if ring.luar_bulan:
        rinci = ", ".join(f"{b}: {v:,}" for b, v in sorted(ring.luar_bulan.items()))
        print(f"             dilewati karena di luar {bulan_arg}: "
              f"{sum(ring.luar_bulan.values()):,} baris ({rinci})")
    if ring.cadangan:
        rinci = ", ".join(f"{k.title()}: {v:,}" for k, v in sorted(ring.cadangan.items()))
        print(f"             tanggal '{(utama or '').title()}' kosong/tidak sah -> tanggal diambil "
              f"dari kolom cadangan: {rinci}")
    if ring.rebate_sumber:
        rinci = ", ".join(f"{k.title()}: {v:,}" for k, v in sorted(ring.rebate_sumber.items()))
        print(f"             Rebate Withdrawal: 'Completed Date' DIABAIKAN, tanggal dari {rinci}"
              + (f" ({ring.rebate_selisih:,} baris 'Completed Date'-nya beda > "
                 f"{tgl.AMBANG_SELISIH_REBATE_HARI} hari dari Completed Time)"
                 if ring.rebate_selisih else ""))
    if ring.konflik:
        n_k = sum(ring.konflik.values())
        print(f"             {n_k:,} baris: 'Completed Date' sah tapi bulannya BEDA dari "
              f"Completed Time + Settlement Date -- "
              + ("Completed Time dipakai" if tgl.UTAMAKAN_WAKTU else
                 "'Completed Date' tetap dipakai (lihat sheet 'Date Check')"))
    if ring.n_tanpa:
        rinci = ", ".join(f"{tgl.NAMA_ALASAN.get(k, k)}: {v:,}"
                          for k, v in sorted(ring.tanpa_tanggal.items()))
        print(f"             TANPA tanggal yang bisa dipakai: {ring.n_tanpa:,} baris ({rinci})"
              + (" -- dilewati" if saring else ""))
    if ring.lewat_ambang():
        pct = ring.n_tanpa / ring.total * 100
        print(f"  !! PERINGATAN: sheet {t}: {ring.n_tanpa:,} dari {ring.total:,} baris "
              f"({pct:.1f}%) TIDAK punya tanggal yang bisa dipakai (batas "
              f"{tgl.AMBANG_TANPA_TANGGAL * 100:.0f}%). Baris itu tidak masuk bulan mana "
              f"pun -- cek ekspornya.")
    if ring.status_geser:
        print(f"  !! PERINGATAN: sheet {t}: {ring.status_geser:,} baris punya kata status "
              f"(success/finish/...) di kolom 'Pay Channel' -- kolom bergeser di ekspor "
              f"sumber. Status tetap dibaca dari kolom 'Status'.")


def salin_baris(ws_d, r, hdr_d, hdr_s, row):
    """Salin satu baris sumber ke baris r template, kolom dicocokkan lewat header."""
    for j, h in enumerate(hdr_d, start=1):
        if not h or norm(h) in DIHITUNG:
            continue
        i = hdr_s.get(kunci_hdr(h))
        if i is not None and i < len(row):
            ws_d.cell(r, j, row[i])


STATUS_SELESAI = ("FINISH", "SUCCESS")     # sama dengan hitung_dw.STATUS_DIPAKAI


def baris_settlement_saja(row, c_settle, c_status, c_pay, saring):
    """KVB --settlement-rows: True kalau 'Settlement Date' baris ini jatuh di bulan
    laporan dan statusnya dipakai (finish/success; sheet tanpa kolom Status = dipakai).

    Baris refuse/cancel TIDAK disalin: kolom Settlement Date tidak memakainya, dan
    kalau ikut tersalin, baris refuse bisa masuk aturan H+1 MTOATD di hitung_dw --
    tab lain harus tetap sama persis dengan tanpa flag ini."""
    if c_settle is None or c_settle >= len(row):
        return False
    d = tanggal(row[c_settle])
    if d is None or (d.year, d.month) != tuple(saring):
        return False
    if c_status is not None:
        st = norm(row[c_status]) if c_status < len(row) else ""
        if not st and c_pay is not None and c_pay < len(row) and \
                norm(row[c_pay]) in tgl.KATA_STATUS:
            st = norm(row[c_pay])           # kolom bergeser (ekspor Rebate)
        if st not in STATUS_SELESAI:
            return False
    return True


def tulis_sheet_ringkasan(wb_out):
    """Sheet 'Date Check': hitungan baris per sheet, untuk hitung_dw (Missing Data)
    dan untuk orang yang memeriksa file hasil."""
    if tgl.SHEET_RINGKASAN in wb_out.sheetnames:
        del wb_out[tgl.SHEET_RINGKASAN]
    ws = wb_out.create_sheet(tgl.SHEET_RINGKASAN)
    ws.append(["Date Check: rows read, used and skipped while filling the template"])
    ws.append(["Sheet", "Category", "Detail", "Rows"])
    for ring in RINGKASAN:
        for baris in ring.baris_ringkasan():
            ws.append(list(baris))
    ws.column_dimensions["B"].width = 58
    ws.column_dimensions["C"].width = 70


def isi_fund_transfer(wb_src, wb_out, saring=None):
    """Salin sheet 'Fund Transfer Table' dari sumber ke template.

    Ini yang dipakai sheet 'Channel Balance' (kolom Fund transfer) dan seluruh
    sheet 'J Wallet (calc)'. Tanpa ini kedua sheet itu kosong.

    saring = (tahun, bulan) -> baris yang tanggal 汇款日期-nya di luar bulan itu
    tidak ikut disalin. Ekspor bulanan selalu punya ekor bulan sebelum/sesudah,
    dan kalau ikut terbawa, saldo J Wallet dan kolom 'Fund transfer' di Payment
    Channel Balance memuat uang bulan lain."""
    nama = cari_sheet(wb_src, {norm(FTT_SHEET), "FUND TRANSFER"})
    if nama is None or FTT_SHEET not in wb_out.sheetnames:
        return None, None
    baris = [list(r) for r in wb_src[nama].iter_rows(max_col=30, values_only=True)]
    # cari baris header: yang memuat kolom batas sisi
    ih = next((i for i, r in enumerate(baris[:10])
               if any(kunci_hdr(v) == kunci_hdr(FTT_BATAS_SISI) for v in r)), None)
    if ih is None:
        print(f"  sheet {FTT_SHEET!r}: baris header (kolom {FTT_BATAS_SISI!r}) tidak ketemu -- dilewati")
        return None, None

    ws_d = wb_out[FTT_SHEET]
    hdr_d = [c.value for c in ws_d[FTT_BARIS_HEADER]]
    peta_d = peta_ftt(hdr_d)
    peta_s = peta_ftt(baris[ih])

    if ws_d.max_row > FTT_BARIS_HEADER:
        ws_d.delete_rows(FTT_BARIS_HEADER + 1, ws_d.max_row - FTT_BARIS_HEADER)

    # kolom tanggal sisi 汇出 di SUMBER, untuk saringan bulan
    j_tgl = next((peta_s[("out", kunci_hdr(k))] for k in FTT_KOLOM_TGL
                  if ("out", kunci_hdr(k)) in peta_s), None)
    if saring and j_tgl is None:
        print(f"  sheet {FTT_SHEET!r}: kolom tanggal "
              f"({' / '.join(FTT_KOLOM_TGL)}) tidak ketemu -- saringan bulan "
              f"TIDAK bisa dipakai di sheet ini, semua baris disalin")

    n, tgl_min, dilewati = 0, None, 0
    for row in baris[ih + 1:]:
        if not any(v not in (None, "") for v in row):
            continue
        if saring and j_tgl is not None:
            d = tanggal(row[j_tgl]) if j_tgl < len(row) else None
            if j_tgl < len(row) and tanggal_tidak_sah(row[j_tgl]):
                TANGGAL_TIDAK_SAH.append((FTT_SHEET, None, row[j_tgl], ""))
            if not d or (d.year, d.month) != saring:
                dilewati += 1
                continue
        n += 1
        for kunci, j_d in peta_d.items():
            j_s = peta_s.get(kunci)
            if j_s is not None and j_s < len(row):
                ws_d.cell(FTT_BARIS_HEADER + n, j_d + 1, row[j_s])
                d = tanggal(row[j_s])
                if d and "日期" in str(kunci[1]) and (tgl_min is None or d < tgl_min):
                    tgl_min = d
    if dilewati:
        print(f"  sheet {FTT_SHEET!r}: {dilewati:,} baris di luar bulan laporan dilewati")
    return n, tgl_min


# Sheet acuan (rate fee + kurs) SELALU diambil dari FILE YANG DIUPLOAD, tidak
# pernah dari template bawaan. Rate bisa berubah kapan saja, jadi tidak ada satu
# pun angka rate/kurs yang "disimpan" di tool ini.
#   - kalau sumber punya sheet bernama tepat 'D&W FEE'      -> itu yang dipakai
#   - kalau sumber cuma punya 'D&W TD FEE' + '-add.'        -> keduanya disalin,
#     hitung_dw.py yang menggabungnya (dan '-add.' menimpa)
#   - kalau sumber tidak punya sama sekali                  -> BERHENTI dengan pesan,
#     BUKAN diam-diam memakai angka lama dari template
# Terbukti perlu 26 Aug 2026: NGN diubah jadi 9% di file user -> dulu hasilnya
# tetap 3,5% karena template bawaan yang menang.
SHEET_KURS = "Xero"
FEE_AWALAN = ("D&W FEE", "D&W TD FEE", "DW FEE")


def salin_sheet_penuh(wb_src, wb_out, nama):
    """Timpa sheet 'nama' di template dengan isi sheet yang sama dari sumber.
    Nilai saja -- format di file hasil ditulis ulang oleh hitung_dw.py."""
    asal = cari_sheet(wb_src, {norm(nama)})
    if asal is None or nama not in wb_out.sheetnames:
        return None
    ws_s, ws_d = wb_src[asal], wb_out[nama]
    baris = [list(r) for r in ws_s.iter_rows(values_only=True)]
    while baris and not any(x not in (None, "") for x in baris[-1]):
        baris.pop()
    if not baris:
        return None
    if ws_d.max_row > 0:
        ws_d.delete_rows(1, ws_d.max_row)
    for i, row in enumerate(baris, start=1):
        for j, v in enumerate(row, start=1):
            if v not in (None, ""):
                ws_d.cell(i, j, v)
    return len(baris) - 1


SHEET_CNY_RATE = "CNY RATE"


def salin_cny_rate(wb_src, wb_out):
    """v1.32: bawa sheet 'CNY RATE' dari sumber ke hasil (dibuat kalau template
    belum punya). Tidak ada di sumber = tidak melakukan apa-apa (DPM)."""
    asal = next((x for x in wb_src.sheetnames
                 if norm(x).replace(" ", "") == "CNYRATE"), None)
    if asal is None:
        return None
    if SHEET_CNY_RATE not in wb_out.sheetnames:
        sisip = (wb_out.sheetnames.index(SHEET_KURS) + 1
                 if SHEET_KURS in wb_out.sheetnames else len(wb_out.sheetnames))
        wb_out.create_sheet(SHEET_CNY_RATE, sisip)
    ws_d = wb_out[SHEET_CNY_RATE]
    if ws_d.max_row > 0:
        ws_d.delete_rows(1, ws_d.max_row)
    baris = [list(r) for r in wb_src[asal].iter_rows(values_only=True)]
    while baris and not any(x not in (None, "") for x in baris[-1]):
        baris.pop()
    for i, row in enumerate(baris, start=1):
        for j, v in enumerate(row, start=1):
            if v not in (None, ""):
                sel = ws_d.cell(i, j, v)
                if isinstance(v, (datetime.date, datetime.datetime)):
                    sel.number_format = "yyyy-mm-dd"
    return max(len(baris) - 1, 0)


def muat_hitung_dw():
    """hitung_dw.py dipakai untuk logika baca saldo pembuka (jangan diduplikasi)."""
    import importlib.util
    f = Path(__file__).resolve().parent / "hitung_dw.py"
    if not f.exists():
        return None
    spec = importlib.util.spec_from_file_location("hd_ob", f)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def isi_opening_balance(src_path, wb_out, tgl_awal, tgl_awal_jw=None):
    """Tulis saldo pembuka ke sheet 'Opening Balance' template.

    Dibaca dari sheet 'J Wallet' + 'Payment Channel Balance' di workbook sumber:
    baris terakhir SEBELUM tanggal pertama data. Sumber dibuka ulang tanpa
    read_only karena kedua fungsi itu menyapu sheet dari awal."""
    if OB_SHEET not in wb_out.sheetnames or tgl_awal is None:
        return None
    hd = muat_hitung_dw()
    if hd is None:
        return None
    wb_ob = openpyxl.load_workbook(src_path, read_only=True, data_only=True)
    try:
        jw, tgl_jw = hd.saldo_awal_jwallet(wb_ob, tgl_awal_jw or tgl_awal)
        ch, tgl_ch = hd.saldo_awal_channel(wb_ob, tgl_awal)
    finally:
        wb_ob.close()

    ws = wb_out[OB_SHEET]
    if ws.max_row > OB_BARIS_HEADER:
        ws.delete_rows(OB_BARIS_HEADER + 1, ws.max_row - OB_BARIS_HEADER)
    r = OB_BARIS_HEADER
    if jw:
        r += 1
        for c, v in enumerate((tgl_jw, "USDT", "J Wallet", jw), start=1):
            sel = ws.cell(r, c, v)
            if c == 1:
                sel.number_format = "yyyy-mm-dd"
    for (cur, ch_nama), v in sorted(ch.items()):
        r += 1
        for c, x in enumerate((tgl_ch, cur, ch_nama, v), start=1):
            sel = ws.cell(r, c, x)
            if c == 1:
                sel.number_format = "yyyy-mm-dd"
    return r - OB_BARIS_HEADER, tgl_ch, tgl_jw


def main():
    ap = argparse.ArgumentParser(description="Isi template D&W dengan data dari workbook sumber")
    ap.add_argument("input")
    ap.add_argument("-t", "--template", default="Template D&W.xlsx")
    ap.add_argument("-o", "--output")
    ap.add_argument("--bulan", metavar="YYYY-MM",
                    help="BULAN LAPORAN: hanya baris bulan ini yang disalin "
                         "(sheet D, W dan Fund Transfer Table)")
    ap.add_argument("--tanpa-extra-fees", action="store_true",
                    help="Saat menggabungkan beberapa fee sheet, jangan tambahkan fallback "
                         "EXTRA_FEES milik DPM (dipakai untuk KVB).")
    ap.add_argument("--settlement-rows", action="store_true",
                    help="KVB (v1.30): dengan --bulan, salin JUGA baris yang tanggal "
                         "transaksinya di luar bulan tapi 'Settlement Date'-nya di dalam "
                         "bulan dan statusnya finish/success. Baris itu HANYA dipakai "
                         "kolom Settlement Date di D&W Report (hitung_dw "
                         "--settlement-columns); saldo pembuka tidak bergeser.")
    ap.add_argument("--sheet-d", help="nama sheet deposit di sumber (default: D / Deposit)")
    ap.add_argument("--sheet-w", help="nama sheet withdrawal di sumber (default: W / Withdrawal)")
    args = ap.parse_args()

    src = Path(args.input).expanduser()
    tpl = Path(args.template).expanduser()
    if not tpl.is_absolute() and not tpl.exists():
        tpl = Path(__file__).resolve().parent / args.template
    if not src.exists():
        sys.exit(f"File sumber tidak ditemukan: {src}")
    if not tpl.exists():
        sys.exit(f"Template tidak ditemukan: {tpl}\n"
                 f"Bikin dulu:  python3 buat_template.py \"{src.name}\"")

    dst = Path(args.output).expanduser() if args.output else \
        src.with_name(f"{src.stem} - siap hitung.xlsx")

    saring = None
    if args.bulan:
        bagian = str(args.bulan).strip().replace("/", "-").split("-")
        if len(bagian) != 2 or not all(x.isdigit() for x in bagian):
            sys.exit(f"--bulan harus berbentuk YYYY-MM (mis. 2026-06), bukan {args.bulan!r}")
        th, bl = int(bagian[0]), int(bagian[1])
        if not 1 <= bl <= 12 or not 2000 <= th <= 2099:
            sys.exit(f"--bulan di luar akal: {args.bulan!r} (bulan 1-12, tahun 2000-2099)")
        saring = (th, bl)
        print(f"Bulan    : {datetime.date(th, bl, 1):%b %Y} "
              f"-- baris di luar bulan ini tidak disalin")

    if tpl.resolve() == src.resolve():
        sys.exit(f"Sumber dan template file-nya SAMA ({tpl.name}).\n"
                 f"Tentukan template lain dengan -t, atau pakai file sumber yang benar.")

    print(f"Template : {tpl.name}")
    print(f"Sumber   : {src.name}")
    wb_out = openpyxl.load_workbook(tpl)
    wb_src = openpyxl.load_workbook(src, read_only=True, data_only=True)

    # Template dasar harus BERSIH. Kalau sheet D/W-nya sudah berisi data, berarti
    # yang dipakai bukan template kosong -- biasanya karena file hasil pernah
    # tersimpan menimpa 'Template D&W.xlsx'. Isi lamanya memang dihapus di bawah,
    # tapi sheet 'Xero' dan 'D&W FEE' TIDAK, jadi kurs/rate-nya bisa ikut cacat
    # (mis. Xero cuma memuat 1 bulan). Peringatkan, jangan diam-diam.
    for cek in ("D", "W"):
        if cek in wb_out.sheetnames:
            n_lama = sum(1 for row in wb_out[cek].iter_rows(min_row=2, values_only=True)
                         if row and row[0] not in (None, ""))
            if n_lama:
                print(f"  !! PERINGATAN: template {tpl.name!r} sheet {cek!r} sudah berisi "
                      f"{n_lama:,} baris data -- ini bukan template kosong.")
                print(f"     Isi itu akan dihapus, TAPI sheet 'Xero' dan 'D&W FEE' di "
                      f"template itu tidak diperiksa.")
                print(f"     Kalau kurs/rate-nya ternyata tidak lengkap, bikin ulang "
                      f"template-nya:  python3 buat_template.py \"workbook-sumber.xlsx\"")
                break
    n_kurs = sum(1 for row in wb_out["Xero"].iter_rows(min_row=2, max_col=1, values_only=True)
                 if row and row[0] not in (None, "")) if "Xero" in wb_out.sheetnames else 0
    if n_kurs < 60:
        print(f"  !! PERINGATAN: sheet 'Xero' di template cuma {n_kurs} tanggal. "
              f"Transaksi di luar rentang itu tidak akan dapat kurs.")

    RINGKASAN.clear()
    total = 0
    tgl_awal = [None]        # tanggal paling awal di data -> basis saldo pembuka
    # Bulan apa saja yang benar-benar ada di file. Dipakai kalau saringan bulan
    # menyapu semuanya: "kamu minta Jan 2026, file ini isinya 2026-05 dan 2026-06"
    # jauh lebih berguna daripada "tidak ada baris data".
    bulan_ada = {}
    # Dicoba berkelompok, dari yang paling spesifik -- supaya 'D' menang atas
    # 'Deposits' walaupun 'Deposits' muncul lebih dulu di urutan sheet.
    for target, opsi, grup in (("D", args.sheet_d, ({"D"}, {"DEPOSIT"}, {"DEPOSITS"})),
                               ("W", args.sheet_w, ({"W"}, {"WITHDRAWAL", "WD"}, {"WITHDRAWALS"}))):
        if target not in wb_out.sheetnames:
            sys.exit(f"Template tidak punya sheet {target!r}")
        nama = opsi
        if nama is None:
            for kandidat in grup:
                nama = cari_sheet(wb_src, kandidat)
                if nama:
                    break
        if nama is None:
            semua = sorted(x for g in grup for x in g)
            print(f"  sheet {target}: sumber tidak punya sheet yang cocok "
                  f"({' / '.join(semua)}) -- dilewati")
            continue
        if nama not in wb_src.sheetnames:
            sys.exit(f"Sheet {nama!r} tidak ada di sumber. Sheet yang ada: {wb_src.sheetnames}")

        ws_s = wb_src[nama]
        baris = [r for r in ws_s.iter_rows(values_only=True)]
        if not baris:
            print(f"  sheet {target}: {nama!r} kosong -- dilewati")
            continue
        hdr_s = {}
        for i, v in enumerate(baris[0]):
            if v not in (None, ""):
                hdr_s.setdefault(kunci_hdr(v), i)
        ws_d = wb_out[target]
        hdr_d = [c.value for c in ws_d[1]]

        # bersihkan isi lama template (header dibiarkan)
        if ws_d.max_row > 1:
            ws_d.delete_rows(2, ws_d.max_row - 1)

        c_cur = hdr_s.get("CURRENCY")
        c_tgl = next((hdr_s[k] for k in KOLOM_TGL if k in hdr_s), None)
        utama = next((k for k in KOLOM_TGL if k in hdr_s), None)
        # Kolom yang ikut menentukan tanggal: kolom utama + cadangannya (tanggal_util).
        kol_tanggal = {k: hdr_s[k] for k in tgl.kolom_penentu(utama) if k in hdr_s}
        c_pay = hdr_s.get("PAY CHANNEL")
        ring = tgl.Ringkasan(target)
        RINGKASAN.append(ring)
        if c_cur is None:
            print(f"  sheet {target}: kolom 'Currency' tidak ketemu di {nama!r} "
                  f"(header terbaca: {', '.join(list(hdr_s)[:8])}) -- dilewati")
            continue

        tak_ada = [h for h in hdr_d
                   if h and kunci_hdr(h) not in hdr_s and norm(h) not in DIHITUNG]
        c_settle = hdr_s.get("SETTLEMENT DATE")
        c_status = hdr_s.get("STATUS")
        n = dilewati = n_settle = 0
        for nomor_baris, row in enumerate(baris[1:], start=2):
            if c_cur is None or c_cur >= len(row) or row[c_cur] in (None, ""):
                continue
            nilai = {k: (row[j] if j < len(row) else None) for k, j in kol_tanggal.items()}
            hasil = (tgl.tentukan_tanggal(nilai, utama) if utama else
                     tgl.HasilTanggal(None, None, "kosong", None, None))
            mentah = nilai.get(utama)
            d = hasil.tanggal
            if c_pay is not None and c_pay < len(row) and \
                    norm(row[c_pay]) in tgl.KATA_STATUS:
                ring.status_geser += 1
            if d is None and hasil.alasan_utama == "tidak_sah":
                info = []
                for k in ("REFERENCE", "TICKET", "PAYMENT GATEWAY", "TRANSACTION",
                          "APPLY DATE", "SETTLEMENT DATE"):
                    j = hdr_s.get(k)
                    if j is not None and j < len(row) and row[j] not in (None, ""):
                        x = row[j]
                        if k.endswith("DATE"):
                            x = tanggal(x) or x
                        info.append(f"{k.title()}={x}")
                TANGGAL_TIDAK_SAH.append((target, nomor_baris, mentah, ", ".join(info)))
            if saring:
                if d is None and hasil.alasan_utama == "tidak_sah":
                    kunci = "(tanggal tidak sah)"
                else:
                    kunci = f"{d:%Y-%m}" if d else "(tanpa tanggal)"
                bulan_ada[kunci] = bulan_ada.get(kunci, 0) + 1
                masuk = bool(d) and (d.year, d.month) == saring
                simpan_settle = (not masuk and d is not None and args.settlement_rows
                                 and baris_settlement_saja(row, c_settle, c_status,
                                                           c_pay, saring))
                ring.catat(hasil, utama, masuk, settle_only=simpan_settle)
                if simpan_settle:
                    # Disalin, tapi TIDAK menggeser tgl_awal (basis saldo pembuka):
                    # hitung_dw memakainya HANYA untuk kolom Settlement Date.
                    n += 1
                    n_settle += 1
                    salin_baris(ws_d, n + 1, hdr_d, hdr_s, row)
                    continue
                if not masuk:
                    dilewati += 1
                    continue
            else:
                ring.catat(hasil, utama, None)
                ring.dipakai += 1
            n += 1
            # Basis saldo pembuka = tanggal TRANSAKSI paling awal (Paid / Completed
            # Date, atau cadangannya), bukan Settlement Date kalau Completed Date ada.
            # Settlement bisa jatuh di bulan sebelumnya, dan itu akan menggeser
            # saldo pembuka sehari terlalu jauh.
            if d and (tgl_awal[0] is None or d < tgl_awal[0]):
                tgl_awal[0] = d
            salin_baris(ws_d, n + 1, hdr_d, hdr_s, row)
        print(f"  sheet {target}: {n:,} baris dari {nama!r}"
              + (f", {dilewati:,} di luar {args.bulan} dilewati" if dilewati else ""))
        if n_settle:
            print(f"             {n_settle:,} di antaranya tanggal transaksinya di luar "
                  f"{args.bulan} tapi Settlement Date-nya di dalam -> HANYA untuk kolom "
                  f"Settlement Date di D&W Report")
        laporkan_ringkasan(ring, utama, saring, args.bulan)
        if tak_ada:
            print(f"             kolom template yang tidak ada di sumber (dibiarkan kosong): "
                  f"{', '.join(tak_ada)}")
        total += n

    # --- kurs
    n = salin_sheet_penuh(wb_src, wb_out, SHEET_KURS)
    if n:
        print(f"  sheet {SHEET_KURS!r}: {n:,} baris kurs disalin dari file yang diupload")
    else:
        sys.exit(f"File itu tidak punya sheet kurs {SHEET_KURS!r}.\n"
                 f"Tool ini TIDAK menyimpan kurs -- kursnya harus ada di file yang "
                 f"kamu upload.\n"
                 f"Tambahkan sheet '{SHEET_KURS}' (baris 1: Date + kode currency, "
                 f"baris berikutnya tanggal + kursnya), lalu coba lagi.")

    # --- CNY RATE (KVB, opsional)
    n = salin_cny_rate(wb_src, wb_out)
    if n is not None:
        print(f"  sheet {SHEET_CNY_RATE!r}: {n:,} baris rate disalin dari file yang diupload")

    # --- tabel fee: apa pun sheet fee yang ada di sumber, itu yang dipakai
    fee_src = [x for x in wb_src.sheetnames
               if any(norm(x).startswith(a) for a in FEE_AWALAN)]
    if not fee_src:
        sys.exit("File itu tidak punya tabel fee.\n"
                 "Tool ini TIDAK menyimpan rate -- tabel fee-nya harus ada di file yang "
                 "kamu upload.\n"
                 "Tambahkan sheet 'D&W FEE' (kolom: Currency, Payment Gateway, Deposit, "
                 "Withdrawal, dan opsional Deposit/Withdrawal Fixed & Min), lalu coba lagi.")
    # buang sheet fee milik template supaya tidak ada angka lama yang ikut
    for x in [x for x in wb_out.sheetnames
              if any(norm(x).startswith(a) for a in FEE_AWALAN)]:
        del wb_out[x]
    for x in fee_src:
        ws_s = wb_src[x]
        ws_d = wb_out.create_sheet(x)
        for row in ws_s.iter_rows(values_only=True):
            ws_d.append(list(row))
        print(f"  sheet {x!r}: {ws_d.max_row:,} baris tabel fee disalin dari file yang diupload")

    # Kalau yang disalin lebih dari satu sheet fee (mis. 'D&W TD FEE' +
    # 'D&W TD FEE -add.'), GABUNG jadi SATU sheet 'D&W FEE' lalu buang yang mentah.
    # Hasil kerjanya harus punya satu tabel fee saja -- itu yang diedit kalau rate
    # berubah, dan kolom Source/Note-nya memperlihatkan angka mana yang menimpa apa.
    if len(fee_src) > 1:
        hd = muat_hitung_dw()
        if hd is not None:
            tabel = hd.load_fee_table(wb_out, pakai_extra_fees=not args.tanpa_extra_fees)
            if tabel:
                # write_fee_sheet membuang semua sheet fee lain lalu menulis satu
                n_fee = hd.write_fee_sheet(wb_out, tabel)
                print(f"  sheet {hd.FEE_SHEET!r}: {n_fee} kombinasi -- hasil GABUNGAN "
                      f"{len(fee_src)} sheet fee ('-add.' menimpa), sheet mentahnya dibuang")

    n_ftt, tgl_ftt = isi_fund_transfer(wb_src, wb_out, saring)
    if n_ftt is not None:
        print(f"  sheet {FTT_SHEET!r}: {n_ftt:,} baris pemindahan dana")
    wb_src.close()

    hasil_ob = isi_opening_balance(src, wb_out, tgl_awal[0], tgl_ftt)
    if hasil_ob:
        n_ob, tgl_ch, tgl_jw = hasil_ob
        print(f"  sheet {OB_SHEET!r}: {n_ob:,} saldo pembuka "
              f"(channel per {tgl_ch}, J Wallet per {tgl_jw})")
    elif OB_SHEET in wb_out.sheetnames:
        print(f"  sheet {OB_SHEET!r}: sumber tidak punya sheet 'J Wallet' / "
              f"'Payment Channel Balance' -> saldo mulai dari nol")

    if RINGKASAN:
        tulis_sheet_ringkasan(wb_out)

    if TANGGAL_TIDAK_SAH:
        print(f"\n  !! PERINGATAN: {len(TANGGAL_TIDAK_SAH)} baris bertanggal di luar "
              f"{TAHUN_MIN}-{TAHUN_MAX} (mis. 1970) -- TIDAK masuk bulan mana pun. "
              f"Perbaiki tanggalnya di sumber kalau baris ini seharusnya ikut:")
        for sh, nb, v, info in TANGGAL_TIDAK_SAH[:100]:
            print(f"     sheet {sh!r}" + (f" baris {nb}" if nb else "")
                  + f": tanggal {v!r}" + (f" ({info})" if info else ""))
        # Catatan verifikasi di file hasil, supaya tidak hilang bersama log.
        if SHEET_TGL_TIDAK_SAH in wb_out.sheetnames:
            del wb_out[SHEET_TGL_TIDAK_SAH]
        ws_v = wb_out.create_sheet(SHEET_TGL_TIDAK_SAH)
        ws_v.append([f"Verifikasi: baris bertanggal di luar {TAHUN_MIN}-{TAHUN_MAX} "
                     f"-- tidak masuk bulan laporan mana pun"])
        ws_v.append(["Sheet", "Row in source", "Date value", "Details"])
        for sh, nb, v, info in TANGGAL_TIDAK_SAH:
            ws_v.append([sh, nb, str(v), info])

    if total == 0:
        if saring and bulan_ada:
            punya = ", ".join(f"{k} ({v:,} baris)"
                              for k, v in sorted(bulan_ada.items()))
            sys.exit(f"BULAN LAPORANNYA SALAH: tidak ada satu pun baris "
                     f"{datetime.date(saring[0], saring[1], 1):%b %Y} di file itu.\n"
                     f"Yang ada di file ini: {punya}.\n"
                     f"Pilih bulan yang benar di halaman upload, atau upload file "
                     f"yang memang memuat bulan itu.")
        sys.exit("TIDAK ADA BARIS DATA di file itu.\n"
                 "Sheet 'D' dan 'W'-nya cuma berisi baris header, jadi tidak ada yang "
                 "bisa dihitung.\n"
                 "Tempel data deposit ke sheet 'D' mulai baris 2 dan withdrawal ke "
                 "sheet 'W' mulai baris 2\n"
                 "(biarkan baris header-nya), simpan, lalu coba lagi.")
    wb_out.save(dst)
    print(f"\nSimpan   : {dst}  ({dst.stat().st_size / 1024:,.0f} KB)")
    print(f"           {total:,} baris siap dihitung")
    print("\nSekarang jalankan kalkulatornya pada file itu.")


if __name__ == "__main__":
    main()
