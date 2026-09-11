#!/usr/bin/env python3
"""PATOKAN UTAMA sejak 7 Sep 2026 -- angka acuan RESMI dari tim Malaysia.

Lebih kuat dari banding_mtoatd_juni.py, karena sheet 'June 26 - MTOATD.xlsx' yang
mereka kirim TERBUKTI SALAH di beberapa sel (mereka sendiri yang mengoreksi).
Sepuluh angka di bawah ini yang mereka nyatakan BENAR.

    python3 cek_acuan_tim.py "file-hasil-juni.xlsx"

Harus keluar COCOK 10/10. Kalau tidak, ada yang rusak.
"""
import datetime
import sys

import openpyxl

F = sys.argv[1] if len(sys.argv) > 1 else "../D&W JUN 2026-hasil.xlsx"
J1 = datetime.date(2026, 6, 1)
J17 = datetime.date(2026, 6, 17)

# (tanggal, currency, label baris, nilai benar, sel di sheet mereka)
ACUAN = [
    (J1,  "VND",  "MT4出金 (USD)",        139_806.19,        "C18"),
    (J1,  "VND",  "Wallet出金 (USD)",       5_652.00,        "C19"),
    (J1,  "VND",  "TD出金 (USD)",         154_295.51,        "C21"),
    (J1,  "VND",  "CRM 出金 （原币种）", 3_782_244_777.00,     "C35"),
    (J1,  "VND",  "实出（原币种）",      4_012_126_167.00,     "C36"),
    (J17, "USDT", "MT4出金 (USD)",         72_558.41,       "C834"),
    (J17, "USDT", "Wallet出金 (USD)",      31_525.00,       "C835"),
    (J17, "USDT", "TD出金 (USD)",          69_080.12,       "C837"),
    (J17, "USDT", "CRM 出金 （原币种）",     103_981.91,       "C850"),
    (J17, "USDT", "实出（原币种）",          68_998.12,       "C852"),
]


def dd(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def baca(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "MTOATD" not in wb.sheetnames:
        sys.exit(f"Sheet 'MTOATD' tidak ada di {path}")
    rows = [list(r) for r in wb["MTOATD"].iter_rows(max_col=24, values_only=True)]
    wb.close()
    out = {}
    for i, r in enumerate(rows):
        if not r or not str(r[0] or "").strip().startswith("日期") or not dd(r[1]):
            continue
        t = dd(r[1])
        kol = {}
        for j in range(2, min(len(r), 24)):
            v = str(r[j] or "").strip()
            if v == "合計":
                break
            if v and not dd(r[j]):
                kol[v] = j
        for k in range(i + 1, min(i + 51, len(rows))):
            lab = str(rows[k][1] or "").strip() if len(rows[k]) > 1 else ""
            if not lab or lab.startswith("日期"):
                continue
            for c, j in kol.items():
                v = rows[k][j] if j < len(rows[k]) else None
                if isinstance(v, (int, float)):
                    out[(lab, t, c)] = float(v)
    return out


A = baca(F)
print(f"File : {F}\n")
print(f"{'sel':7} {'tanggal':12} {'cur':6} {'baris':24} {'hasil kita':>20} {'acuan tim':>20}")
print("-" * 96)
ok = 0
for t, c, lab, tgt, sel in ACUAN:
    x = A.get((lab, t, c))
    cocok = x is not None and abs(x - tgt) < 0.01
    ok += cocok
    tampil = f"{x:,.2f}" if x is not None else "(tidak ada)"
    print(f"{sel:7} {str(t):12} {c:6} {lab:24} {tampil:>20} {tgt:>20,.2f}"
          f"  {'OK' if cocok else '<-- BEDA'}")
print("-" * 96)
print(f"COCOK {ok}/{len(ACUAN)}")
if ok != len(ACUAN):
    print("\nADA YANG RUSAK. Aturan yang harus berlaku:")
    print("  1. Baris （原币种）SELALU kolom P (Transaction), termasuk USDT & USD.")
    print("  2. Baris berbasis SETTLEMENT DATE (MT4出金 / Wallet出金 / CRM出金（原币种）)")
    print("     ikut menghitung baris refuse yang Completed Date = Settlement Date + 1 hari.")
    print("  3. Baris berbasis COMPLETED DATE (TD出金 / 实出（原币种）) tetap membuang refuse.")
    sys.exit(1)
