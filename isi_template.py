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


def tanggal(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


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
        if c_cur is None:
            print(f"  sheet {target}: kolom 'Currency' tidak ketemu di {nama!r} "
                  f"(header terbaca: {', '.join(list(hdr_s)[:8])}) -- dilewati")
            continue

        tak_ada = [h for h in hdr_d
                   if h and kunci_hdr(h) not in hdr_s and norm(h) not in DIHITUNG]
        n = dilewati = 0
        for row in baris[1:]:
            if c_cur is None or c_cur >= len(row) or row[c_cur] in (None, ""):
                continue
            if saring:
                d = tanggal(row[c_tgl]) if c_tgl is not None and c_tgl < len(row) else None
                kunci = f"{d:%Y-%m}" if d else "(tanpa tanggal)"
                bulan_ada[kunci] = bulan_ada.get(kunci, 0) + 1
                if not d or (d.year, d.month) != saring:
                    dilewati += 1
                    continue
            n += 1
            # Basis saldo pembuka = tanggal TRANSAKSI paling awal (Paid / Completed
            # Date), bukan Settlement Date. Settlement bisa jatuh di bulan
            # sebelumnya, dan itu akan menggeser saldo pembuka sehari terlalu jauh.
            d = tanggal(row[c_tgl]) if c_tgl is not None and c_tgl < len(row) else None
            if d and (tgl_awal[0] is None or d < tgl_awal[0]):
                tgl_awal[0] = d
            for j, h in enumerate(hdr_d, start=1):
                if not h or norm(h) in DIHITUNG:
                    continue
                i = hdr_s.get(kunci_hdr(h))
                if i is not None and i < len(row):
                    ws_d.cell(n + 1, j, row[i])
        print(f"  sheet {target}: {n:,} baris dari {nama!r}"
              + (f", {dilewati:,} di luar {args.bulan} dilewati" if dilewati else ""))
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
            tabel = hd.load_fee_table(wb_out)
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
