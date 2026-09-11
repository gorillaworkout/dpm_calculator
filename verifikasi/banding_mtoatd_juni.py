#!/usr/bin/env python3
"""Banding sheet 'MTOATD' hasil kita vs MTOATD JUNI 2026 asli tim Malaysia.

PATOKAN BARU (3 Sep 2026) -- jauh lebih kuat dari banding_mtoatd.py yang cuma
1 hari: acuan ini SEBULAN PENUH, 30 blok tanggal.

    python3 banding_mtoatd_juni.py "file-hasil.xlsx" ["June 26 - MTOATD.xlsx"]

Hasil yang HARUS keluar (file hasil dari 'June - ..._Test 1.xlsx', --period 2026-06):

    COCOK 11.259   BEDA 1.341   (8 Sep 2026; sebelumnya 11.291/1.309)

PERHATIAN: sheet 'June 26 - MTOATD.xlsx' SUDAH TIDAK JADI PATOKAN UTAMA. Tim
Malaysia sendiri mengoreksi beberapa selnya (mis. Wallet出金 VND 1 Jun tertulis
6.455,00 padahal yang benar 5.652,00). Patokan utama sekarang
`cek_acuan_tim.py` -- sepuluh angka yang mereka nyatakan BENAR, harus 10/10.
Skrip ini tetap berguna untuk melihat POLA selisih, bukan untuk lulus/gagal.

    Sisi DEPOSIT           : 0 beda  <- MT4入金, TD入金, 实收, CRM入金（原币种）
    MT4出金 (USD)           : 0 beda  <- BUKTI aturan 'hanya finish' kita BENAR:
                                        rumus mereka tidak menyaring status, tapi
                                        ekspor Withdrawals mereka memang sudah
                                        tidak memuat baris refuse.
    Wallet出金 (rebate)     : 157 beda, total USD -163.828  <- BEDA DATA, bukan rumus.
    CRM出金 / 本日净入金       : selisihnya PERSIS selisih Wallet出金 (570/570), warisan.
    TD出金                  : -163.885 = -163.828 (rebate) + -57,00 (baris status KOSONG)

Jadi beda ATURAN sebulan penuh cuma USD 57,00 dari 83 baris yang kolom Status-nya
kosong: SUMIFS "<>refuse" mereka ikut menghitung sel kosong, kita tidak.
Sisanya murni data rebate yang tidak ada di sheet 'W' kita.

CATATAN 3 Sep 2026: sempat diubah supaya baris （原币种）selalu kolom P untuk
semua currency (patokan jadi 11.393/1.207), lalu DIKEMBALIKAN setelah tim
Malaysia menegaskan kolom M untuk USDT/USD memang disengaja. Lihat
OC_KOLOM_USD di mtoatd_spec.py.

Kalau 'BEDA' NAIK jauh di atas 1.108, atau daftar COCOK SEMPURNA turun dari
12 baris (sisi DEPOSIT + MT4出金 + ——手续费/转账费), berarti ada yang rusak.
"""
import datetime
import os
import sys

import openpyxl

HASIL = sys.argv[1] if len(sys.argv) > 1 else "../D&W JUN 2026-hasil.xlsx"
ACUAN = sys.argv[2] if len(sys.argv) > 2 else os.path.expanduser(
    "~/Downloads/June 26 - MTOATD.xlsx")
TOL = 0.02


def dd(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def baca(path, sheet):
    """-> {(label, tanggal, currency): nilai}. Satu blok = 51 baris, label di kolom B."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        sys.exit(f"Sheet {sheet!r} tidak ada di {path}\nYang ada: {wb.sheetnames}")
    rows = [list(r) for r in wb[sheet].iter_rows(max_col=24, values_only=True)]
    wb.close()
    out = {}
    for i, r in enumerate(rows):
        # blok pertama pakai '日期：' (titik dua lebar), sisanya '日期:' -- cocokkan awalannya
        if not r or not str(r[0] or "").strip().startswith("日期") or not dd(r[1]):
            continue
        t = dd(r[1])
        kol = {}
        for j in range(2, min(len(r), 24)):
            v = str(r[j] or "").strip()
            if v == "合計":
                break                       # sesudah ini blok akumulasi, bukan currency
            if v and not dd(r[j]):
                kol[v] = j
        for k in range(i + 1, min(i + 51, len(rows))):
            lab = str(rows[k][1] or "").strip() if len(rows[k]) > 1 else ""
            if not lab or lab.startswith("日期"):
                continue
            for c, j in kol.items():
                v = rows[k][j] if j < len(rows[k]) else None
                out[(lab, t, c)] = float(v) if isinstance(v, (int, float)) else None
    return out


A = baca(HASIL, "MTOATD")
B = baca(ACUAN, "MTOATD (MAY'26)")
tgl_a = {t for _l, t, _c in A}
tgl_b = {t for _l, t, _c in B}
irisan = sorted(tgl_a & tgl_b)
print(f"blok kami {len(tgl_a)}  |  blok acuan {len(tgl_b)}  |  beririsan {len(irisan)}")
if not irisan:
    sys.exit("Tidak ada tanggal yang sama -- pastikan file hasilnya bulan Juni 2026.")

cocok = beda = kosong = 0
per_beda = {}
per_cocok = {}
total = {}
contoh = {}
for (lab, t, c), y in B.items():
    if t not in irisan or y is None:
        continue
    x = A.get((lab, t, c))
    if x is None:
        kosong += 1
        continue
    if abs(x - y) < TOL:
        cocok += 1
        per_cocok[lab] = per_cocok.get(lab, 0) + 1
    else:
        beda += 1
        per_beda[lab] = per_beda.get(lab, 0) + 1
        total[lab] = total.get(lab, 0.0) + (x - y)
        contoh.setdefault(lab, []).append((t, c, round(x, 2), round(y, 2)))

print(f"\nCOCOK {cocok:,}   BEDA {beda:,}   (kami sengaja kosong: {kosong:,})")
print("\nPATOKAN 8 Sep 2026: COCOK 11.259, BEDA 1.341 -> NORMAL.")
print("Sheet acuan ini punya sel yang SALAH (dikoreksi tim sendiri).")
print("Patokan lulus/gagal yang sebenarnya: cek_acuan_tim.py -> harus 10/10.")
print("Sisi deposit dan MT4出金 HARUS 0 beda. Kalau tidak, ada yang rusak.\n")
print(f"{'BARIS':34} {'BEDA':>6} {'COCOK':>7} {'total selisih (kami-mereka)':>28}")
for lab in sorted(set(per_beda) | set(per_cocok), key=lambda x: -per_beda.get(x, 0)):
    if not per_beda.get(lab):
        continue
    print(f"{lab:34} {per_beda[lab]:>6} {per_cocok.get(lab, 0):>7} {total.get(lab, 0.0):>28,.2f}")
    for c in contoh[lab][:2]:
        print(f"     {c[0]}  {c[1]:5} kami {c[2]:>18,.2f}   mereka {c[3]:>18,.2f}")

bersih = [lab for lab in per_cocok if not per_beda.get(lab)]
print(f"\nCOCOK SEMPURNA ({len(bersih)} baris): {', '.join(sorted(bersih))}")
