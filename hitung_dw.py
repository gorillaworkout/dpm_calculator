#!/usr/bin/env python3
"""
Auto-hitung Handling Fee, Xero Rate, Xero USD, dan Forex Gain/Loss
untuk workbook D&W (Deposit & Withdrawal) Dupoin Markets.

Pakai:
    python3 hitung_dw.py "DPM-D&W.xlsx"                     -> tulis ke DPM-D&W-hasil.xlsx
    python3 hitung_dw.py "DPM-D&W.xlsx" -o "Juni.xlsx"      -> nama output sendiri
    python3 hitung_dw.py "DPM-D&W.xlsx" --in-place           -> timpa file aslinya
    python3 hitung_dw.py "DPM-D&W.xlsx" --sheet D W          -> pilih sheet manual
    python3 hitung_dw.py "export-baru.xlsx" --fee-from "DPM-D&W.xlsx" --xero-from "DPM-D&W.xlsx"
        -> file data baru yg belum punya sheet 'D&W FEE' / 'XERO', ambil dari file lama

Yang dihitung (4 kolom baru di kanan sheet D / W):
    Handling Fee     = Transaction x rate(dari sheet 'D&W FEE', kolom Deposit/Withdrawal)
    Xero Rate        = kurs di sheet 'XERO' pada Paid Date + Currency
    Xero USD         = Transaction / Xero Rate      (lihat XERO_RATE_MULTIPLIER)
    Forex Gain/Loss  = Xero USD - USD (kolom USD yang sudah ada)

Kalau skrip dijalankan ulang, ke-4 kolom itu ditimpa (tidak menumpuk).

Sheet hasil yang ditulis -- SEKALI JALAN, tidak ada tahap lanjutan lagi:
    D&W Report        ringkasan Currency x Payment Gateway
    D&W Detail        gabungan D + W, satu baris per transaksi
    D&W FEE           tabel fee tergabung + catatan rate mana yang menimpa apa
    Channel Balance   Date x Currency x Channel + saldo berjalan
    J Wallet (calc)   buku besar dompet USDT dari Fund Transfer Table
    MTOATD            rekonsiliasi harian MT4 / Wallet / CRM / TD  (rumus: mtoatd_spec.py)
    Missing Data      data yang kurang / perlu dikonfirmasi
    Legend            penjelasan kolom & warna (English)

Butuh: pip install openpyxl
"""

import argparse
import datetime
import os
import re
import threading
import time
import sys
from collections import Counter, defaultdict
from pathlib import Path

# Console Windows default-nya cp1252/cp437 yang tidak punya karakter seperti '（'
# (dipakai di nama gateway 'PA-PHP（Manual）'), jadi print() bisa crash.
# Paksa UTF-8; kalau tetap tidak bisa, karakter aneh diganti '?' daripada error.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.comments import Comment
except ImportError:
    if sys.platform.startswith("win"):
        sys.exit("openpyxl belum terpasang.\n"
                 "Jalankan di Command Prompt:\n"
                 "    py -m pip install openpyxl\n"
                 "(pakai 'py -m pip', bukan 'pip' -- 'pip' sering belum masuk PATH)")
    sys.exit("openpyxl belum terpasang. Jalankan:\n"
             "    python3 -m pip install --user openpyxl")

# ==========================================================================
# KONFIGURASI  -- bagian ini yang biasanya perlu kamu edit
# ==========================================================================

# Rate yang belum ada di sheet 'D&W FEE' -> (CURRENCY, GATEWAY): (deposit, withdrawal)
EXTRA_FEES = {
    #                        deposit  withdrawal
    ("VND", "NOVOLINK"):    (0.009,   0.000),    # ditambahkan 19 Aug 2026
    ("USD", "BUZIPAY"):     (0.000,   0.000),    # ditambahkan 19 Aug 2026
    ("THB", "VPAY"):        (0.025,   0.002),    # ditambahkan 19 Aug 2026
    ("THB", "NEPAY"):       (0.027,   0.005),    # ditambahkan 19 Aug 2026
    ("THB", "BERRY"):       (0.024,   0.000),    # ditambahkan 19 Aug 2026
    ("UZS", "MNTX"):        (0.055,   0.035),    # ditambahkan 19 Aug 2026 (khusus UZS)
    ("USD", "MNTX-LAK USD"): (None,   0.000),    # 19 Aug 2026: isi kolom Withdrawal yg kosong di sheet
    # Nilainya boleh berupa teks juga, mis. ("NGN","MNTX"): ("1.5%+50", "3.5%+55")
    ("USD", "PAY247-KH USD"): (None,  0.020),    # 19 Aug 2026: isi kolom Withdrawal yg kosong di sheet
    ("VND", "BECKPAY"):     (0.016,   0.000),    # ditambahkan 19 Aug 2026
}

# Nama gateway di sheet data -> nama gateway di tabel fee (semua di-UPPERCASE)
GATEWAY_ALIAS = {
    ("VND", "VN77PAY"):         "77 PAY",
    ("USDT", "USDT-BEP20"):     "OTHER",        # BEP20 tdk terdaftar; semua USDT 0%
    ("USDT", "BEP20"):          "OTHER",        # penulisan di sheet W
    ("USD", "MNTX-BINANCEPAY"): "BINANCE PAY",
    ("USD", "MNTX"):            "MNTX-USD",     # dikonfirmasi 21 Aug 2026:
                                                # 'MNTX' di bawah USD memang MNTX-USD
}

# True  = Xero Rate disimpan sebagai multiplier (1/kurs), Xero USD = Transaction x Xero Rate
# False = Xero Rate disimpan apa adanya spt di sheet XERO, Xero USD = Transaction / Xero Rate
XERO_RATE_MULTIPLIER = False

# Tambahkan fixed fee (kolom ke-5 tabel fee, mis. VND/77 Pay = 10.000/trx) ke Handling Fee?
INCLUDE_FIXED_FEE = False

# Kolom tanggal yang dipakai untuk cari kurs di sheet XERO, urut prioritas.
# Sheet 'D' punya 'Paid Date'; sheet 'W' tidak punya, jadi turun ke 'Completed Date'.
DATE_COLUMNS = ("PAID DATE", "COMPLETED DATE", "SETTLEMENT DATE", "APPLY DATE")

# HANYA baris berstatus 'finish' yang dihitung (aturan user, 28 Aug 2026).
# Withdrawal yang ditolak tidak pernah dibayarkan, jadi tidak boleh masuk ke
# laporan mana pun -- bukan cuma dikecualikan dari handling fee.
# CATATAN: sebagian baris gagal kolom Status-nya KOSONG, bukan 'refuse'
# (30 Jun 2026: 139 refuse + 26 kosong, remark-nya sama-sama "Application Failed").
# Karena aturannya "hanya finish", yang kosong ikut dikeluarkan juga.
# Sheet yang tidak punya kolom Status sama sekali (mis. 'D') tidak terpengaruh.
HANYA_STATUS_FINISH = True
STATUS_DIPAKAI = ("FINISH",)

# ------------------------------------------------------------- periode laporan
# Bulan laporan, diisi dari --period YYYY-MM  ->  (tahun, bulan). None = semua
# tanggal dipakai (perilaku lama).
#
# KENAPA PERLU: ekspor back office selalu punya EKOR. File "Juni" yang diupload
# user berisi 25 Mei s/d 5 Jul, karena yang disaring waktu ekspor adalah Apply
# Date sementara laporan memakai Paid Date (deposit) / Completed Date
# (withdrawal). Tanpa saringan ini sheet MTOATD keluar 42 blok tanggal, bukan 30,
# dan totalnya memuat uang bulan lain.
#
# Basis tanggalnya SAMA dengan yang dipakai seluruh laporan: kolom `Date` per
# baris = kolom tanggal yang dipilih per sheet (DATE_COLUMNS, urut prioritas) --
# 'Paid Date' untuk D, 'Completed Date' untuk W. Jadi sebuah baris masuk bulan
# yang sama di D&W Report, MTOATD, dan Payment Channel Balance.
#
# Baris yang tanggalnya KOSONG ikut dibuang saat periode dipilih: tanpa tanggal
# baris itu tidak bisa ditempatkan di bulan mana pun. Jumlahnya dilaporkan
# terpisah ("(no date)") di Terminal dan di sheet 'Missing Data'.
PERIODE_FILTER = None


def set_periode(teks):
    """'2026-06' -> (2026, 6). Dipakai --period. String kosong/None = tanpa saringan."""
    global PERIODE_FILTER
    if not teks:
        PERIODE_FILTER = None
        return None
    t = str(teks).strip().replace("/", "-")
    bagian = t.split("-")
    if len(bagian) != 2 or not bagian[0].isdigit() or not bagian[1].isdigit():
        sys.exit(f"--period harus berbentuk YYYY-MM (mis. 2026-06), bukan {teks!r}")
    th, bl = int(bagian[0]), int(bagian[1])
    if not 1 <= bl <= 12 or not 2000 <= th <= 2099:
        sys.exit(f"--period di luar akal: {teks!r} (bulan 1-12, tahun 2000-2099)")
    PERIODE_FILTER = (th, bl)
    return PERIODE_FILTER


# Saldo pembuka J Wallet yang DIKETIK MANUSIA di halaman upload (--jwallet-opening).
# Ini menang atas sheet 'J Wallet' maupun 'Opening Balance'. Gunanya: file yang
# diupload sering TIDAK memuat sheet saldo apa pun, jadi J Wallet mulai dari NOL.
# Dengan angka ini, saldo Juni bisa diteruskan dari saldo penutup Mei.
SALDO_AWAL_JW = None


def baca_angka_saldo(teks):
    """'377233.36' / '377.233,36' / '377,233.36' -> 377233.36. None kalau kosong.

    Pemisah ribuan vs desimal ditebak: kalau ada DUA jenis pemisah, yang PALING
    KANAN itu desimal. Kalau cuma satu jenis, dianggap desimal hanya bila diikuti
    1-2 digit ('377,23' -> 377.23); selain itu pemisah ribuan ('377.233' -> 377233).
    """
    if teks is None:
        return None
    t = str(teks).strip().replace(" ", "").replace("_", "")
    if not t:
        return None
    neg = t.startswith("-")
    t = t.lstrip("+-")
    if not t or not all(c.isdigit() or c in ".," for c in t):
        sys.exit(f"--jwallet-opening bukan angka yang bisa dibaca: {teks!r}")
    if "." in t and "," in t:
        des = "." if t.rfind(".") > t.rfind(",") else ","
        t = t.replace("," if des == "." else ".", "").replace(des, ".")
    elif "." in t or "," in t:
        sep = "." if "." in t else ","
        ekor = t.rsplit(sep, 1)[1]
        t = t.replace(sep, ".") if (t.count(sep) == 1 and 1 <= len(ekor) <= 2) \
            else t.replace(sep, "")
    try:
        v = float(t)
    except ValueError:
        sys.exit(f"--jwallet-opening bukan angka yang bisa dibaca: {teks!r}")
    return -v if neg else v


def tanggal_saldo_pembuka():
    """Tanggal yang dicatat untuk saldo pembuka: hari terakhir bulan SEBELUM
    bulan laporan. Juni 2026 -> 31 Mei 2026. Tanpa periode -> None."""
    if PERIODE_FILTER is None:
        return None
    th, bl = PERIODE_FILTER
    return datetime.date(th, bl, 1) - datetime.timedelta(days=1)


def dalam_periode(d):
    """True kalau tanggal d masuk bulan laporan. Tanpa periode -> selalu True."""
    if PERIODE_FILTER is None:
        return True
    return bool(d) and (d.year, d.month) == PERIODE_FILTER


def nama_periode():
    """'Jun 2026' untuk pesan ke manusia. Tanpa periode -> ''."""
    if PERIODE_FILTER is None:
        return ""
    return datetime.date(PERIODE_FILTER[0], PERIODE_FILTER[1], 1).strftime("%b %Y")

# Tag di nama gateway yang diabaikan saat cari rate, mis. 'PA-PHP（Manual）' -> 'PA-PHP'
# (（） adalah tanda kurung full-width yang dipakai di data back office)
GATEWAY_STRIP_TAGS = ("（MANUAL）", "(MANUAL)", "（MANUAL", "MANUAL）")

# Kalau Paid Date tidak ada di sheet XERO: "nearest" = pakai tanggal terdekat sebelumnya,
# "exact" = kosongkan dan laporkan sebagai error.
MISSING_DATE_MODE = "nearest"

# Kalau gateway tidak ketemu di tabel fee, jatuh ke baris "<CURRENCY>/Other"?
# True  = hanya kalau nama gateway di data memang literal "Other".
#         Gateway lain yang tidak ketemu -> Handling Fee DIKOSONGKAN + ditandai MERAH,
#         supaya tidak diam-diam dihitung 0%.
# False = tidak pernah pakai baris Other sama sekali.
FALLBACK_TO_OTHER = True
LITERAL_OTHER_NAMES = ("OTHER", "OTHERS", "LAIN-LAIN", "LAINNYA")

# Forex Gain/Loss untuk baris withdrawal dibalik tandanya?
#   False -> withdrawal ikut rumus deposit : Xero USD - USD
#   True  -> withdrawal dibalik           : USD - Xero USD   (uang keluar)
FLIP_FOREX_ON_WITHDRAWAL = True

# Cetak progress tiap berapa baris. Penting: file besar butuh menit-menit,
# tanpa ini prosesnya terlihat "stuck".
PROGRESS_EVERY = 2000

# Rate yang ANGKANYA MASIH DIPERTANYAKAN. Baris ini tetap dipakai untuk menghitung,
# tapi ditandai KUNING di sheet 'D&W FEE' + dilaporkan di sheet 'Missing Data',
# supaya orang yang membaca file tahu angkanya belum final.
# Aturan yang dipakai (ditetapkan user 24 Aug 2026): kalau 'D&W TD FEE' dan
# 'D&W TD FEE -add.' berbeda, yang dipakai '-add.'. Dari 25 kombinasi yang ada di
# kedua sheet, 10 berbeda. Aturan itu TERBUKTI BENAR untuk THB/12PAY (2,7%),
# VND/77PAY (2,5%) dan KRW/MNTX (4%) -- dicek terhadap 'Payment Channel Balance'
# mereka sendiri (rate nyata = Deposit charges / Deposit).
# TAPI untuk tiga kombinasi di bawah, data mereka menyalahkan aturan itu.
# TIDAK ADA RATE YANG DISIMPAN DI KODE (26 Aug 2026, permintaan user).
# Rate bisa berubah kapan saja, jadi satu-satunya sumber adalah sheet fee di FILE
# YANG DIUPLOAD. Dulu ada `RATE_KEPUTUSAN` yang menimpa hasil gabungan untuk
# memperbaiki NGN/MNTX yang tertukar di sheet '-add.' -- itu sudah DIHAPUS.
#
# Gantinya cuma PEMERIKSAAN, bukan angka: kalau tabel fee memuat kombinasi yang
# sudah pernah terbukti tertukar, baris itu ditandai supaya diperbaiki di Excel-nya.
# Format: (currency, gateway): (deposit_yang_salah, withdrawal_yang_salah, penjelasan)
CURIGA_TERTUKAR = {
    ("NGN", "MNTX"): (0.015, 0.035,
                      "This looks like the swapped pair from the '-add.' sheet. "
                      "Verified against 'Payment Channel Balance': the correct values are "
                      "deposit 3.5% + 50 per transaction (fits 114 of 115 days) and "
                      "withdrawal 1.5% + 55 per transaction (fits 132 of 135). "
                      "Please correct this row in the 'D&W FEE' sheet of your file."),
    ("EGP", "MNTX"): (0.07, 0.035,
                      "Deposit 7% / withdrawal 3.5% is the main-table orientation. "
                      "Confirmed 26 Aug 2026 that deposit should be 3.5%. "
                      "Please correct this row in the 'D&W FEE' sheet of your file."),
}

PERLU_KONFIRMASI = {
    # Kosong: rate yang dipertanyakan sudah dijawab, dan sekarang semua rate
    # dibaca dari Excel-nya. Lihat CURIGA_TERTUKAR di atas dan CATATAN_RATE di bawah.
}

# Catatan informasi (bukan peringatan) untuk baris rate tertentu di sheet 'D&W FEE'.
CATATAN_RATE = {
    ("NGN", "MNTX"): ("CONFIRMED 26 Aug 2026: deposit 3.5% + 50 per transaction, "
                      "withdrawal 1.5% + 55 per transaction. The two fee tables had the "
                      "percentages swapped - the '-add.' table said deposit 1.5% / "
                      "withdrawal 3.5%, which is wrong on both sides. Verified against your "
                      "own 'Payment Channel Balance': deposit 3.5% fits 114 of the 115 days "
                      "on record, withdrawal 1.5% fits 132 of 135, while the swapped version "
                      "fits 32/115 and 0/135. Proof for the fixed amount: on three different "
                      "days the deposit total was the same 140,000 but the charge was "
                      "4,950 / 5,000 / 5,050 = 140,000 x 3.5% plus 50 for EACH transaction "
                      "that day. So the fixed fee is per transaction, not once a day."),
    ("PHP", "SHUNFA PAY"): ("CONFIRMED 26 Aug 2026: withdrawal is the fixed 8 per "
                            "TRANSACTION, no percentage - so 10 withdrawals on 1 Jun 2026 = "
                            "80.00. Your 'Payment Channel Balance' shows 200.00 for that "
                            "same day and channel (= 20 x 10), which is incorrect. "
                            "Deposit is 1.8% per the '-add.' table; note that sheet shows "
                            "2.1000% on 39 of the 40 days on record."),
    ("EGP", "MNTX"): ("CONFIRMED AGAIN 3 Sep 2026: the team was asked directly whether the "
                      "deposit rate is the 2.25% written into their 'Payment Channel "
                      "Balance' formula or the 3.5% in the fee table, and answered "
                      "\"Follow the fee table\". So 3.5% it is - which is what this report "
                      "already uses. (Earlier: confirmed 26 Aug 2026, deposit 3.5% / "
                      "withdrawal 7%, the '-add.' table winning as usual.) Note the actual "
                      "rate implied by 'Payment Channel Balance' is not constant - 7.0000% "
                      "on 13 Jan and 2 Feb 2026, around 2.07-2.42% on other dates - so a "
                      "deposit charge here may not tie back exactly."),
    ("THB", "12PAY"): ("Verified from your own 'Payment Channel Balance': exactly "
                       "2.7000% on all 153 days on record. The '-add.' value (2.7%) is "
                       "correct, the main table's 2.8% is superseded."),
    ("KRW", "MNTX"): ("Verified from your own 'Payment Channel Balance': exactly "
                      "4.0000% on all 16 days on record. The '-add.' value (4%) is "
                      "correct, the main table's 5% is superseded."),
    ("USD", "MNTX"): ("Confirmed 21 Aug 2026: 'MNTX' under USD means MNTX-USD. "
                      "Both resolve to deposit 3% / withdrawal 2%, so the figures are "
                      "the same either way."),
    ("VND", "77PAY"): ("Rate changed over time: 2.7% in Dec 2025, moved during Jan 2026, "
                       "2.5% from Feb 2026 onwards. Verified from the data (Deposit "
                       "charges / Deposit) - 2.5% fits 122 of the 150 days on record, "
                       "and every day from Feb 2026. The '-add.' value (2.5%) is the "
                       "current rate; the main table's 2.7% is the old one."),
}

# Penjelasan untuk kolom yang belum ada rumusnya -- dipasang sebagai komentar
# di sel header supaya orang yang mengisi tahu apa yang dicari.
INFO_WITHDRAWAL_CHARGES = (
    "SOLVED 26 Aug 2026 - this column is the sum of the withdrawal Handling Fee, i.e. "
    "withdrawal amount x the withdrawal rate from the fee table.\n\n"
    "That is exactly what your own 'Payment Channel Balance' does, wherever the cell "
    "actually has a formula. Two styles are used there:\n"
    "  = SUMIFS(Withdrawals!$Y:$Y, ...)   -> sums the 'Handling Fee' column, same as us\n"
    "  = <withdrawal cell> * 1.5%         -> amount x a rate typed into the formula\n\n"
    "Channels where we used to differ turned out to have NO FORMULA AT ALL in that "
    "sheet - the number is typed in by hand. On 1 June 2026 that is the case for "
    "VND 77 PAY, VND OC PAY, NGN MONETIX, PHP PA, PHP SHUNFA PAY, INR Pay247-INR, "
    "THB 1-2 Pay and every USDT channel. So the old '24,159,169.51 unexplained excess' "
    "on 77 PAY was never computed from anything - there is no formula to reproduce.\n\n"
    "One case proves it cannot be a rate: INR / Pay247-INR on 1 June 2026 has "
    "Withdrawal = 0 but Withdrawal charges = 12,709.78.\n\n"
    "ANSWERED 3 Sep 2026 by the DPM Malaysia team: \"Continue to use the same formula, "
    "our side will check with other reports and edit those that need to be edited "
    "manually.\"\n\n"
    "So the value in this column IS the agreed calculation - withdrawal amount x the "
    "withdrawal rate from the fee table - and it is not in question. The cell stays "
    "YELLOW for one reason only: this is the column the finance team may still adjust by "
    "hand against gateway statements after the report is produced. Treat it as the "
    "calculated starting point, not as an error."
)

NEW_COLUMNS = ["Handling Fee", "Xero Rate", "Xero USD", "Forex Gain/Loss"]

# Nama kolom hitungan yang sudah dipakai di file lama. Kalau salah satu ditemukan,
# kolom ITU yang diisi dan judulnya TIDAK diubah -- jadi tampilan file tetap sama.
NEW_COLUMN_ALIASES = {
    "Handling Fee":    ["HANDLING FEE", "HANDLING CHARGES"],
    "Xero Rate":       ["XERO RATE"],
    "Xero USD":        ["XERO USD"],
    "Forex Gain/Loss": ["FOREX GAIN/LOSS", "FOREX GAIN/-LOSS", "FOREX GAIN / LOSS",
                        "CURRENCY GAIN/LOSS", "CURRENCY GAIN/-LOSS", "FE GAIN/-LOSS"],
}

# Header cell comments (English -- this file is shared with other people)
LEGENDA = {
    "Handling Fee":
        "Handling Fee = Transaction x fee rate from the 'D&W FEE' sheet,\n"
        "matched on Currency + Payment Gateway.\n"
        "Deposit rows use the Deposit rate, withdrawal rows the Withdrawal rate.\n\n"
        "RED CELL = no fee rate exists yet for this Currency + Payment Gateway.\n"
        "The cell is deliberately left EMPTY (not 0) so the amount is never\n"
        "silently understated.\n"
        "TO FIX: add the missing row to the 'D&W FEE' sheet, then run again.",
    "Xero Rate":
        "Exchange rate taken from the 'XERO' sheet, looked up by\n"
        "transaction date + currency. Quoted as local currency per 1 USD.\n\n"
        "YELLOW CELL = that date is not in the XERO sheet yet, so the rate of\n"
        "the nearest earlier date was used instead.\n"
        "TO FIX: add the missing dates to the 'XERO' sheet, then run again.\n\n"
        "RED CELL = no rate found at all for this date and currency.",
    "Xero USD":
        "Xero USD = Transaction / Xero Rate\n"
        "The transaction amount restated in USD at the XERO rate.\n\n"
        "RED CELL = could not be calculated (no exchange rate).",
    "CRM Average Rate":
        "TRANSACTIONS / CRM USD for that Currency + Payment Gateway.\n\n"
        "On the Grand Total row this is the same ratio computed on the totals.\n"
        "It mixes currencies (THB, VND, USD, ...) so treat the Grand Total\n"
        "average as a control figure, not an exchange rate.",
    "CRM average":
        "Transactions / CRM USD for that Currency + Payment Gateway.\n\n"
        "On the Grand Total row this is the same ratio computed on the totals.\n"
        "It mixes currencies, so treat it as a control figure only.",
    "Data Status":
        "Per-row data quality note.\n\n"
        "OK   = fee rate and exchange rate both found exactly.\n"
        "FEE: = no fee rate for this Currency + Payment Gateway.\n"
        "FX:  = the transaction date is missing from the XERO sheet, or no rate\n"
        "       exists for that currency at all.\n"
        "DATA:= the source export itself is incomplete (empty gateway, amount\n"
        "       or date) and must be fixed in the back office.\n\n"
        "See the 'Missing Data' sheet for the full list and what to add.",
    "Date":
        "The date used to look up the exchange rate.\n"
        "See the 'Date Source' column for which original column it came from\n"
        "(deposit rows use Paid Date, withdrawal rows use Completed Date).",
    "Forex Gain/Loss":
        "The difference between the USD value at the XERO rate and the USD\n"
        "amount recorded in the source data.\n\n"
        "Deposit rows    : Xero USD - USD\n"
        "Withdrawal rows : USD - Xero USD   (sign flipped: money going out)\n\n"
        "RED CELL = could not be calculated (no exchange rate).",
}

# Sheet legenda (bahasa Inggris) -- dibuat/diperbarui tiap run
LEGEND_SHEET = "Legend"

# Sheet hasil gabungan D + W, dan sheet daftar data yang kurang.
# Dibuat ulang tiap run. Matikan dengan --no-report kalau tidak perlu.
# --- Payment Channel Balance & J Wallet ---------------------------------------
# Sumber pemindahan dana. Header dicari otomatis (biasanya baris 3).
FTT_SHEET_NAMA = ("FUND TRANSFER TABLE", "FUND TRANSFER")
# Nama sheet hasil dibuat SAMA dengan workbook asli mereka (permintaan user
# 28 Aug 2026) supaya file hasil bisa langsung menggantikan yang lama.
# LAYOUT-nya tetap yang baru (rapi, Date x Currency x Channel) -- bukan layout
# lebar ~500 kolom punya mereka.
# HATI-HATI: nama ini BENTROK dengan sheet input yang kita baca saldo pembukanya.
# Karena itu saldo pembuka HARUS dibaca DULU sebelum sheet lama dihapus --
# lihat write_channel_sheets().
PCB_SHEET = "Payment Channel Balance"   # hasil: Date x Currency x Channel, layout tidy
JW_SHEET = "J Wallet"                   # hasil: buku besar dompet + saldo berjalan
MAKE_CHANNEL_SHEETS = True

# Rumus saldo channel (terverifikasi 11.611/12.127 baris di file Juni 2026):
#   Balance(t) = Balance(t-1) + Deposit - Deposit charges - Fund transfer
#                - Fund transfer charges - Withdrawal - Withdrawal charges
# 'Withdrawal charges' IKUT dikurangkan (dibetulkan 26 Aug 2026). Dulu tidak,
# karena kolomnya masih kosong; sekarang kolomnya terisi dan saldo mereka memang
# menguranginya -- terbukti: tiap selisih saldo persis sebesar Withdrawal charges.
# 'Withdrawal charges' = jumlah withdrawal x rate withdrawal dari tabel fee.
# Sudah DIISI, tapi tetap ditandai KUNING: di beberapa channel sheet mereka punya
# kelebihan yang belum terjelaskan (lihat INFO_WITHDRAWAL_CHARGES).
# Gabungkan 'VN77PAY（Manual）' ke 'VN77PAY' di sheet 'Channel Balance'?
#   False (default) = dipisah, SAMA seperti sheet mereka.
#   True            = digabung jadi satu baris.
# Dicek 26 Aug 2026: sheet 'Payment Channel Balance' mereka TIDAK punya kolom
# （Manual）sama sekali, dan angka 364.175.000 tidak muncul di mana pun di sheet itu.
# Withdrawal '77 PAY' 1 Jun 2026 di sheet mereka = 1.081.158.848 -- persis nilai
# VN77PAY TANPA baris Manual, dan saldo penutup mereka (3.464.768.698) cocok sampai
# sen dengan angka itu. Jadi payout manual memang TIDAK lewat rekening gateway,
# sehingga tidak menggerakkan saldo channel. Menggabungkan = selisih 364.175.000.
GABUNG_TAG_MANUAL = False

PCB_KOLOM = ["Date", "Currency", "Payment Gateway", "Deposit", "Deposit charges",
             "Fund transfer", "Fund transfer charges", "Withdrawal",
             "Withdrawal charges", "Balance", "Status"]

MTOATD_SHEET = "MTOATD"           # rekonsiliasi harian, dihitung penuh dari D + W
MAKE_MTOATD_SHEET = True

# Tulis juga blok 当月累计 (akumulasi sejak awal bulan) di sebelah kanan,
# seperti sheet MTOATD yang sudah beredar.
MTOATD_BLOK_AKUMULASI = True

# Rumus MTOATD ada di mtoatd_spec.py. Semuanya dibongkar dari sheet
# 'MTOATD (MAY''26)' di workbook sumber dan terverifikasi 523/523 sel.
# Baris tanpa Payment Gateway TIDAK dikeluarkan: rumus asli mereka juga tidak
# menyaring gateway, jadi angkanya cocok apa adanya.

# Urutan tab di file hasil, dari paling KIRI (permintaan user 28 Aug 2026).
# Nama yang tidak ada di file hasil dilewati; sheet lain ditaruh di belakang
# dengan urutan apa adanya.
URUTAN_SHEET = [
    "J Wallet",
    "Fund Transfer Table",
    "Payment Channel Balance",
    "MTOATD",
    "D&W Report",
    "D&W Detail",
    "D",
    "W",
    "Xero",
    "D&W TD FEE",
    "Opening Balance",
    "Guide",
    "Missing Data",
    "Legend",
]

REPORT_SHEET = "D&W Report"       # ringkasan pivot: Currency x Payment Gateway
DETAIL_SHEET = "D&W Detail"       # tabel datar gabungan D + W (baris per transaksi)

# Judul & label sheet 'D&W Report' -- disamakan dgn report yang sudah beredar
REPORT_TITLE = "Realised Foreign Exchange Rate Gain/ Loss for Deposit and Withdrawal"
REPORT_SECTION_TITLE = {"deposit": "Deposit [CRM vs Xero]",
                        "withdrawal": "Withdrawal [CRM vs Xero]"}

# Urutan grup di report. Pivot lama urutannya tersimpan di cache Excel dan tidak
# bisa direproduksi, jadi pilih sendiri:
#   "value"      = nilai CRM USD terbesar dulu (default)
#   "appearance" = urutan kemunculan di sheet D / W
#   "alpha"      = A-Z
REPORT_GROUP_ORDER = "value"

# Warna & format, diambil dari report yang sudah ada
RPT_FILL_HEAD = "ACB9CA"      # header kolom data
RPT_FILL_HEAD2 = "A9D08E"     # header kolom hitungan
RPT_FILL_CUR = "DDEBF7"       # sel Currency
RPT_FILL_CALC = "E2EFDA"      # sel kolom hitungan
FMT_ACC = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '
FMT_ACC5 = '_ * #,##0.00000_ ;_ * \\-#,##0.00000_ ;_ * "-"??_ ;_ @_ '
MISSING_SHEET = "Missing Data"
MAKE_REPORT_SHEET = True

# Kolom sheet 'D&W Report', urutannya mengikuti report D&W yang sudah beredar.
# Tiap entri: (nama kolom di report, daftar nama header yang dicari di sheet D/W)
REPORT_COLUMNS = [
    ("Source Name",     ["SOURCE NAME"]),
    ("Sales UID",       ["SALES UID"]),
    ("Sales Name",      ["SALES NAME"]),
    ("Sales Dept",      ["SALES DEPT"]),
    ("Sales Area",      ["SALES AREA"]),
    ("User ID",         ["USER ID"]),
    ("Account",         ["ACCOUNT"]),
    ("Date",            None),          # tanggal yang dipakai untuk cari kurs
    ("Time",            ["PAID TIME", "COMPLETED TIME", "APPLY TIME"]),
    ("Paid Date",       ["PAID DATE"]),         # MTOATD: basis TD入金 / 实收
    ("Completed Date",  ["COMPLETED DATE"]),    # MTOATD: basis TD出金 / 实出
    ("Settlement Date", ["SETTLEMENT DATE"]),   # MTOATD: basis MT4 / CRM
    ("Apply Date",      ["APPLY DATE"]),
    ("Currency",        ["CURRENCY"]),
    ("MT Order",        ["MT ORDER"]),
    ("Ticket",          ["TICKET"]),
    ("USD",             ["USD"]),
    ("Rate",            ["RATE"]),
    ("Transaction",     ["TRANSACTION"]),
    ("Charges",         ["CHARGES"]),           # MTOATD: baris ——手续费/转账费
    ("Payment Gateway", ["PAYMENT GATEWAY"]),
    ("Reference",       ["REFERENCE"]),
    ("Hash",            ["HASH"]),
    ("Status",          ["STATUS"]),
    ("Remark",          ["REMARK", "REMARKS", "NOTE"]),   # alasan kalau ditolak
]

# Kolom audit di ujung sheet report. False = report berisi kolom data saja.
REPORT_AUDIT_COLUMNS = True

# Nama sheet dicocokkan tanpa peduli huruf besar/kecil dan spasi berlebih.
# Sheet fee: semua sheet yang namanya dimulai salah satu ini akan DIGABUNG,
# dan yang mengandung "-ADD" diterapkan terakhir sehingga menimpa yang lain.
FEE_SHEET_AWALAN = ("D&W FEE", "D&W TD FEE", "DW FEE", "D&W TD FEE -ADD")
XERO_SHEET_NAMA = ("XERO",)

FEE_SHEET, XERO_SHEET = "D&W TD FEE", "XERO"

# Tulis hasil gabungan tabel fee jadi sheet 'D&W FEE' di file hasil?
# True  = file hasil punya SATU tabel fee yang sudah tergabung, lengkap dengan
#         kolom Source & Note supaya kelihatan angka mana yang menimpa apa.
#         Kalau sheet 'D&W FEE' sudah ada, sheet ITU yang dipakai (tidak digabung
#         ulang), jadi hasilnya stabil kalau dijalankan berkali-kali.
# False = tabel fee tetap digabung di memori saja, file tidak ditambahi sheet.
WRITE_FEE_SHEET = True

# Nama sheet data yang dikenali langsung. Sheet lain tetap dipakai kalau
# header-nya memuat REQUIRED_COLUMNS (jadi 'Deposit' / 'Withdrawal' / nama lain ikut kebaca).
# Dicoba berkelompok, dari yang paling spesifik. Kelompok pertama yang ketemu
# dipakai, kelompok berikutnya diabaikan -- supaya di workbook yang punya
# 'D', 'W', DAN 'Withdrawal', 'Withdrawal (2)', dst, hanya 'D' + 'W' yang dihitung.
DATA_SHEET_GROUPS = (
    ("D", "W"),
    ("DEPOSIT", "WITHDRAWAL", "WD"),
    ("DEPOSIT DATA", "WITHDRAWAL DATA"),
)
DATA_SHEET_NAMES = tuple(n for g in DATA_SHEET_GROUPS for n in g)
REQUIRED_COLUMNS = ("CURRENCY", "TRANSACTION", "PAYMENT GATEWAY", "USD")

CURRENCY_SUFFIXES = ("VND", "PHP", "USD", "USDT", "INR", "THB", "KHR", "LAK",
                     "MXN", "NGN", "BRL", "ZAR", "EGP", "TRY", "KRW", "JPY", "AED",
                     "UZS", "PKR", "KES", "SGD", "USDC")

# Format angka di Excel. Tanda # = tampilkan kalau perlu, 0 = selalu tampilkan.
# Jadi "#,##0.##" -> 8,100 (bukan 8,100.00) dan 0.18 tetap 0.18
FMT_FEE   = "#,##0.##"        # Handling Fee    : tanpa .00 kalau bulat
FMT_RATE  = "#,##0.######"    # Xero Rate       : 26,317.2 / 61.5507 / 1
FMT_MONEY = "#,##0.00"        # Xero USD & Forex: tetap 2 desimal (nilai uang USD)
FILL_WARN = PatternFill("solid", fgColor="FFF2CC")     # kuning = kurs tanggal terdekat dipakai
FILL_MISSING = PatternFill("solid", fgColor="FFC7CE")  # merah  = butuh data, belum bisa dihitung
FILL_SKIP = PatternFill("solid", fgColor="F2F2F2")     # abu-abu = memang tidak kena fee (ditolak)
FONT_MISSING = Font(color="9C0006", bold=True)

# ==========================================================================


_T0 = time.monotonic()


def jam():
    d = time.monotonic() - _T0
    return f"[{int(d) // 60:d}:{int(d) % 60:02d}]"


def lapor(pesan, ganti_baris=False):
    """Cetak progress. ganti_baris=True menimpa baris yang sama (progress hidup)."""
    if ganti_baris and sys.stdout.isatty():
        print(f"\r{jam()} {pesan}".ljust(78), end="", flush=True)
    else:
        if ganti_baris:
            return
        print(f"{jam()} {pesan}", flush=True)


class Denyut:
    """Penghitung detik yang berjalan selama tahap lambat (buka / simpan file),
    supaya jelas prosesnya masih hidup dan bukan hang.

    Dipakai sebagai context manager:
        with Denyut("membuka file"):
            wb = openpyxl.load_workbook(src)
    """

    PUTAR = "|/-\\"

    def __init__(self, pesan, tenang=False):
        self.pesan = pesan
        self.tenang = tenang or not sys.stdout.isatty()
        self._stop = threading.Event()
        self._th = None

    def _jalan(self):
        i = 0
        mulai = time.monotonic()
        while not self._stop.wait(0.25):
            i += 1
            if i % 4:                       # perbarui tampilan 1x per detik
                continue
            d = int(time.monotonic() - mulai)
            print(f"\r{jam()} {self.pesan} {self.PUTAR[(i // 4) % 4]} {d}s"
                  .ljust(78), end="", flush=True)

    def __enter__(self):
        if self.tenang:
            print(f"{jam()} {self.pesan}...", flush=True)
        else:
            print(f"\r{jam()} {self.pesan} | 0s".ljust(78), end="", flush=True)
            self._th = threading.Thread(target=self._jalan, daemon=True)
            self._th.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        if self._th:
            self._th.join(timeout=1)
            print("\r".ljust(80) + "\r", end="", flush=True)
        return False


def norm(v):
    return " ".join(str(v if v is not None else "").strip().upper().split())


def kunci_gw(v):
    """Kunci pencocokan gateway: buang semua yang bukan huruf/angka.
    Bikin '77 Pay' == '77Pay', '1-2-PAY' == '1-2Pay', 'SHUNFA PAY' == 'SHUNFAPAY'.
    Tanpa ini tabel fee utama dan tabel '-add.' tidak saling menimpa karena
    gaya penulisan nama gateway-nya berbeda."""
    return re.sub(r"[^A-Z0-9]", "", norm(v))


def to_float(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.datetime.strptime(str(v).strip(), fmt).date()
        except (ValueError, TypeError):
            pass
    return None


# ---------------------------------------------------------------- fee table
def cari_sheet_fee(wb):
    """Sheet fee yang dipakai.

    Kalau ada sheet bernama tepat 'D&W FEE' (hasil gabungan dari run sebelumnya),
    HANYA itu yang dipakai. Kalau tidak, semua sheet fee digabung dan yang
    mengandung '-add' diterapkan terakhir sehingga menimpa."""
    tepat = [x for x in wb.sheetnames if norm(x) == norm(FEE_SHEET)]
    if tepat:
        return tepat[:1]
    ketemu = [x for x in wb.sheetnames
              if any(norm(x).startswith(a) for a in FEE_SHEET_AWALAN)]
    return sorted(ketemu, key=lambda x: ("-ADD" in norm(x), x))


def cari_sheet(wb, kandidat):
    for x in wb.sheetnames:
        if norm(x) in kandidat:
            return x
    return None


def baca_rate(v):
    """Rate bisa berupa angka atau teks. -> {'pct','fixed','min'} atau None.

        0.027          -> 2,7%
        1.5%+50        -> 1,5% + biaya tetap 50
        +8             -> 0% + biaya tetap 8
        0.45%,min 90   -> 0,45% dengan minimum 90
    """
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return {"pct": float(v), "fixed": 0.0, "min": None}

    t = str(v).strip()
    if not t:
        return None
    pct = tetap = minimum = None

    m = re.search(r"MIN\w*\s*([\d.,]+)", t, re.I)
    if m:
        minimum = to_float(m.group(1))
        t = t[:m.start()] + t[m.end():]
    m = re.search(r"([\d.]+)\s*%", t)
    if m:
        pct = to_float(m.group(1))
        pct = pct / 100 if pct is not None else None
        t = t[:m.start()] + t[m.end():]
    m = re.search(r"\+\s*([\d.,]+)", t)
    if m:
        tetap = to_float(m.group(1))

    if pct is None and tetap is None and minimum is None:
        f = to_float(t)
        return None if f is None else {"pct": f, "fixed": 0.0, "min": None}
    return {"pct": pct or 0.0, "fixed": tetap or 0.0, "min": minimum}


def _baca_satu_sheet_fee(ws):
    """Parse satu sheet fee. Baris header dan urutan kolom dideteksi dari
    nama header-nya, jadi layout 'Currency dulu' maupun 'Payment Gateway dulu'
    dua-duanya kebaca."""
    baris = list(ws.iter_rows(max_row=min(ws.max_row, 4000), values_only=True))
    h_idx = kol = None
    for i, row in enumerate(baris[:8]):
        nama = {norm(v): j for j, v in enumerate(row) if v not in (None, "")}
        if "CURRENCY" in nama and "PAYMENT GATEWAY" in nama:
            h_idx, kol = i, nama
            break
    if kol is None:
        return {}, None

    c_cur, c_gw = kol["CURRENCY"], kol["PAYMENT GATEWAY"]
    c_dep = kol.get("DEPOSIT")
    # 'WITHDRAWAL' saja -- jangan sampai kena 'WITHDRAWAL FIXED' / 'WITHDRAWAL MIN'
    c_wd = kol.get("WITHDRAWAL")
    if c_wd is None:
        c_wd = next((kol[k] for k in kol
                     if k.startswith("WITHDRAW")
                     and not k.endswith(("FIXED", "MIN"))), None)
    # kolom tambahan yang kita sendiri tulis di sheet 'D&W FEE' hasil gabungan
    c_extra = {
        "dep_fixed": kol.get("DEPOSIT FIXED"),
        "wd_fixed": kol.get("WITHDRAWAL FIXED"),
        "dep_min": kol.get("DEPOSIT MIN"),
        "wd_min": kol.get("WITHDRAWAL MIN"),
    }

    out = {}
    for row in baris[h_idx + 1:]:
        if len(row) <= max(c_cur, c_gw):
            continue
        cur, gw = norm(row[c_cur]), norm(row[c_gw])
        if not cur or not gw:
            continue
        k_gw = kunci_gw(gw)
        dep = baca_rate(row[c_dep]) if c_dep is not None and len(row) > c_dep else None
        wd = baca_rate(row[c_wd]) if c_wd is not None and len(row) > c_wd else None

        # gabungkan kembali komponen tetap / minimum dari kolom terpisah,
        # supaya sheet 'D&W FEE' yang kita tulis bisa dibaca ulang utuh
        def _pakai(komp, kunci_fix, kunci_min):
            c1, c2 = c_extra[kunci_fix], c_extra[kunci_min]
            f = to_float(row[c1]) if c1 is not None and len(row) > c1 else None
            m = to_float(row[c2]) if c2 is not None and len(row) > c2 else None
            if komp is None and (f or m is not None):
                komp = {"pct": 0.0, "fixed": 0.0, "min": None}
            if komp is not None:
                if f:
                    komp["fixed"] = f
                if m is not None:
                    komp["min"] = m
            return komp

        dep = _pakai(dep, "dep_fixed", "dep_min")
        wd = _pakai(wd, "wd_fixed", "wd_min")
        # notasi lama '+10000' di kolom tambahan sheet utama
        legacy = 0.0
        for x in row[max(c_dep or 0, c_wd or 0) + 1:][:4]:
            if isinstance(x, str) and x.strip().startswith("+"):
                legacy = to_float(x.strip().lstrip("+")) or 0.0
        out[(cur, k_gw)] = {"deposit": dep, "withdrawal": wd, "fixed": legacy,
                            "nama": gw}
    return out, h_idx + 1


def load_fee_table(wb):
    """Gabungan semua sheet fee. Sheet '-add.' diterapkan terakhir -> menimpa.

    SUMBER SATU-SATUNYA adalah sheet fee di workbook yang diproses. Tidak ada rate
    yang disimpan di kode. Kalau ada sheet bernama TEPAT 'D&W FEE', itu saja yang
    dipakai; kalau tidak, semua sheet fee digabung dan '-add.' menimpa.
    Jadi kalau bulan depan rate-nya berubah, cukup edit Excel-nya."""
    lembar = cari_sheet_fee(wb)
    table = {}
    for nama in lembar:
        bagian, hrow = _baca_satu_sheet_fee(wb[nama])
        if not bagian:
            print(f"         sheet fee {nama!r}: header 'Currency'+'Payment Gateway' "
                  f"tidak ketemu, dilewati")
            continue
        baru = timpa = 0
        for k, v in bagian.items():
            if k in table:
                timpa += 1
            else:
                baru += 1
            lama = table.get(k, {})
            jejak = list(lama.get("asal", []))
            jejak.append(nama)
            catat = list(lama.get("ditimpa", []))
            for sisi, label in (("deposit", "Deposit"), ("withdrawal", "Withdrawal")):
                x, y = lama.get(sisi), v[sisi]
                if x is not None and y is not None and (
                        abs(x["pct"] - y["pct"]) > 1e-12
                        or abs(x["fixed"] - y["fixed"]) > 1e-12
                        or x["min"] != y["min"]):
                    catat.append(f"{label} di {lama.get('asal', ['?'])[-1]}: "
                                 f"{x['pct'] * 100:g}%"
                                 + (f"+{x['fixed']:g}" if x["fixed"] else "")
                                 + (f" min {x['min']:g}" if x["min"] is not None else ""))
            table[k] = {
                "deposit": v["deposit"] if v["deposit"] is not None else lama.get("deposit"),
                "withdrawal": v["withdrawal"] if v["withdrawal"] is not None else lama.get("withdrawal"),
                "fixed": v["fixed"] or lama.get("fixed", 0.0),
                "nama": v.get("nama") or lama.get("nama", ""),
                "asal": jejak,
                "ditimpa": catat,
            }
        print(f"         sheet fee {nama!r} (header baris {hrow - 1}): "
              f"{len(bagian)} baris -> {baru} baru, {timpa} menimpa")

    # Tidak ada rate yang ditimpa dari kode. Hanya diperiksa: kalau ada kombinasi
    # yang nilainya sama dengan pasangan yang sudah terbukti TERTUKAR, tandai supaya
    # diperbaiki di Excel-nya. Angkanya sendiri dipakai apa adanya.
    for (cur_x, gw_x), (dep_x, wd_x, pesan) in CURIGA_TERTUKAR.items():
        row = table.get((norm(cur_x), kunci_gw(gw_x)))
        if not row:
            continue
        d, w = row.get("deposit"), row.get("withdrawal")
        if d and w and abs(d["pct"] - dep_x) < 1e-12 and abs(w["pct"] - wd_x) < 1e-12:
            row["curiga"] = pesan
            print(f"      !! {cur_x}/{gw_x}: deposit {dep_x*100:g}% / withdrawal "
                  f"{wd_x*100:g}% -- kelihatannya TERTUKAR, betulkan di sheet fee")

    # EXTRA_FEES hanya MENAMBAL yang masih kosong
    for (cur_x, gw_x), (d, w) in EXTRA_FEES.items():
        key = (norm(cur_x), kunci_gw(gw_x))
        row = table.setdefault(key, {"deposit": None, "withdrawal": None, "fixed": 0.0,
                                     "nama": norm(gw_x), "asal": ["EXTRA_FEES (script)"],
                                     "ditimpa": []})
        if row["deposit"] is None and d is not None:
            row["deposit"] = baca_rate(d)
        if row["withdrawal"] is None and w is not None:
            row["withdrawal"] = baca_rate(w)
    return table


def strip_currency_suffix(gw, currency):
    """'BECKPAY-VND' -> 'BECKPAY';  'PA-PHP' -> 'PA'"""
    for suf in (norm(currency),) + CURRENCY_SUFFIXES:
        for sep in ("-", " "):
            tail = sep + suf
            if gw.endswith(tail) and len(gw) > len(tail):
                return gw[: -len(tail)].strip()
    return gw


def strip_tags(gw):
    """'PA-PHP（MANUAL）' -> 'PA-PHP'"""
    for tag in GATEWAY_STRIP_TAGS:
        gw = gw.replace(tag, "")
    return " ".join(gw.replace("（", " ").replace("）", " ").split()).strip(" -")


def lookup_fee(table, currency, gateway, kind):
    """-> (komponen, legacy_fixed, matched_key, how)
    komponen = {'pct','fixed','min'} atau None kalau tidak ketemu."""
    cur, gw = norm(currency), strip_tags(norm(gateway))
    cands = [(gw, "exact"),
             (GATEWAY_ALIAS.get((cur, gw), ""), "alias"),
             (strip_currency_suffix(gw, cur), "strip-suffix")]
    # baris "<CUR>/Other" hanya dipakai kalau gateway-nya memang bernama "Other".
    # Gateway lain yang tidak ketemu sengaja dibiarkan kosong + merah.
    if FALLBACK_TO_OTHER and gw in LITERAL_OTHER_NAMES:
        cands.append(("OTHER", "literal-other"))
    for cand, how in cands:
        if not cand:
            continue
        hit = table.get((cur, kunci_gw(cand)))
        if hit and hit[kind] is not None:
            return hit[kind], hit["fixed"], (cur, hit.get("nama") or cand), how
    return None, 0.0, None, "NOT FOUND"


# --------------------------------------------------------------- xero rates
def load_xero(wb):
    """-> ({(date, CURRENCY): rate}, sorted_dates)"""
    nama = cari_sheet(wb, XERO_SHEET_NAMA)
    if nama is None:
        return {}, []
    ws = wb[nama]
    header = [norm(c.value) for c in ws[1]]
    rates, dates = {}, set()
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        d = as_date(row[0])
        if not d:
            continue
        dates.add(d)
        for col in range(1, min(len(header), len(row))):
            cur, val = header[col], to_float(row[col])
            if cur and val:
                rates[(d, cur)] = val
    return rates, sorted(dates)


def lookup_rate(rates, dates, date, currency):
    """-> (rate, used_date, is_fallback)"""
    cur = norm(currency)
    if date is None:
        return None, None, False
    if (date, cur) in rates:
        return rates[(date, cur)], date, False
    if MISSING_DATE_MODE == "nearest":
        earlier = [d for d in dates if d <= date and (d, cur) in rates]
        if earlier:
            d = earlier[-1]
            return rates[(d, cur)], d, True
        later = [d for d in dates if d > date and (d, cur) in rates]
        if later:
            d = later[0]
            return rates[(d, cur)], d, True
    return None, None, False


# ------------------------------------------------------------------ sheet D/W
# Excel sering mencatat dimensi sheet jauh lebih besar dari data sebenarnya
# (mis. max_row 1.000.000 padahal data 45.000 baris). Berhenti setelah sebanyak
# ini baris kosong berurutan supaya tidak memproses ratusan ribu baris hampa.
BLANK_RUN_STOP = 500


def last_data_row(ws, kolom_kunci):
    """Baris terakhir yang benar-benar berisi data, dilihat dari kolom kunci."""
    lo = min(kolom_kunci)
    hi = max(kolom_kunci)
    idx = [k - lo for k in kolom_kunci]
    last = 1
    blank = 0
    for r, nilai in enumerate(
            ws.iter_rows(min_row=2, min_col=lo, max_col=hi, values_only=True), start=2):
        if any(nilai[i] not in (None, "") for i in idx):
            last = r
            blank = 0
        else:
            blank += 1
            if blank >= BLANK_RUN_STOP:
                break
    return last


def find_columns(ws):
    """-> {header_ternormalisasi: index kolom 1-based} dari baris 1"""
    return {norm(c.value): c.column for c in ws[1] if c.value is not None}


# Isi kolom 'Source Name' yang dianggap withdrawal / deposit (sudah di-UPPERCASE)
WITHDRAWAL_WORDS = ("WITHDRAW", "WD", "W", "OUT", "PENARIKAN", "TARIK", "PAYOUT", "DEBIT")
DEPOSIT_WORDS = ("DEPOSIT", "DP", "D", "IN", "SETOR", "TOPUP", "TOP UP", "CREDIT")


def is_data_sheet(ws):
    """Sheet ini sheet transaksi? Dilihat dari header-nya, bukan namanya."""
    cols = find_columns(ws)
    if not all(n in cols for n in REQUIRED_COLUMNS):
        return False
    return any(n in cols for n in DATE_COLUMNS)


def kind_of(sheet_name, source_name):
    """deposit / withdrawal -- utamakan kolom 'Source Name', fallback nama sheet."""
    s = norm(source_name)
    if s:
        if s in WITHDRAWAL_WORDS or s.startswith("WITHDRAW"):
            return "withdrawal"
        if s in DEPOSIT_WORDS or s.startswith("DEPOSIT"):
            return "deposit"
    n = norm(sheet_name)
    if n in ("W", "WD", "WITHDRAWAL") or n.startswith("WITHDRAW"):
        return "withdrawal"
    return "deposit"


def refuse_h1_relevan(kind, status, completed, settle):
    """Withdrawal refuse yang memengaruhi MTOATD bulan laporan lewat aturan H+1."""
    return (
        kind == "withdrawal"
        and norm(status) == "REFUSE"
        and settle is not None
        and completed == settle + datetime.timedelta(days=1)
        and dalam_periode(settle)
    )


def kosongkan_kolom_hitung(ws, row, out, catatan=None):
    """Kosongkan 4 kolom rumus di satu baris dan tandai ABU-ABU.

    Dipakai untuk baris yang DIBUANG (di luar bulan laporan). Selnya harus benar
    benar dikosongkan, bukan dibiarkan: kalau file hasil diproses ulang dengan
    bulan yang berbeda, angka bulan lama masih menempel di situ dan orang yang
    memfilter sheet D/W di Excel akan menjumlahkan angka yang sudah tidak dipakai.

    JEBAKAN openpyxl: ws.cell(r, c, value=None) TIDAK menghapus isi sel -- None
    diartikan "tidak ada nilai yang diberikan". Harus lewat .value.
    """
    for nama in NEW_COLUMNS:
        kol = out.get(nama)
        if kol is None:
            continue
        cell = ws.cell(row=row, column=kol)
        cell.value = None
        cell.fill = FILL_SKIP           # ABU-ABU = sengaja tidak dihitung
        cell.font = Font()
        if catatan:
            cell.comment = None
    return


def process_sheet(ws, fee_table, rates, xero_dates, report):
    sub = report["per_sheet"][ws.title]
    cols = find_columns(ws)
    missing = [n for n in REQUIRED_COLUMNS if n not in cols]

    # kolom tanggal: ambil yang pertama tersedia sesuai prioritas DATE_COLUMNS
    date_col_name = next((n for n in DATE_COLUMNS if n in cols), None)
    if date_col_name is None:
        missing.append("salah satu dari " + " / ".join(DATE_COLUMNS))

    if missing:
        msg = f"[{ws.title}] kolom wajib tidak ada: {', '.join(missing)}"
        print("  !! " + msg)
        print(f"     header yang terbaca: {', '.join(sorted(cols))}")
        report["errors"].append(msg)
        return 0, None

    c_cur, c_tx = cols["CURRENCY"], cols["TRANSACTION"]
    c_gw, c_date, c_usd = cols["PAYMENT GATEWAY"], cols[date_col_name], cols["USD"]
    c_src = cols.get("SOURCE NAME")
    c_charges = cols.get("CHARGES")
    c_status = cols.get("STATUS")
    c_settle = cols.get("SETTLEMENT DATE")

    # tulis (atau timpa) header 4 kolom baru
    out = {}
    next_col = ws.max_column + 1
    for name in NEW_COLUMNS:
        lama = next((cols[a] for a in NEW_COLUMN_ALIASES[name] if a in cols), None)
        if lama is not None:
            out[name] = lama              # pakai kolom yang sudah ada, judul dibiarkan
            hcell = ws.cell(row=1, column=lama)
        else:
            out[name] = next_col
            next_col += 1
            hcell = ws.cell(row=1, column=out[name], value=name)
            hcell.font = Font(bold=True)
        if name in LEGENDA:
            hcell.comment = Comment(LEGENDA[name], "hitung_dw.py", height=150, width=320)

    with Denyut(f"sheet '{ws.title}': mencari baris data terakhir"):
        baris_akhir = last_data_row(ws, (c_cur, c_gw, c_tx))
    if baris_akhir < 2:
        return 0, date_col_name

    lapor(f"  sheet '{ws.title}': {baris_akhir - 1:,} baris, mulai menghitung...")
    n = 0
    for row in range(2, baris_akhir + 1):
        if row % PROGRESS_EVERY == 0:
            lapor(f"  sheet '{ws.title}': {row - 1:,}/{baris_akhir - 1:,} baris "
                  f"({(row - 1) * 100 // max(baris_akhir - 1, 1)}%)", ganti_baris=True)
        currency = ws.cell(row=row, column=c_cur).value
        gateway = ws.cell(row=row, column=c_gw).value
        if currency is None and gateway is None:
            continue                          # baris kosong
        n += 1

        tx = to_float(ws.cell(row=row, column=c_tx).value)
        usd_existing = to_float(ws.cell(row=row, column=c_usd).value)
        paid = as_date(ws.cell(row=row, column=c_date).value)
        source = ws.cell(row=row, column=c_src).value if c_src else None
        kind = kind_of(ws.title, source)
        status = ws.cell(row=row, column=c_status).value if c_status else None
        settle = as_date(ws.cell(row=row, column=c_settle).value) if c_settle else None

        # Capture sebelum saringan periode: Settlement 30 Jun + Completed 1 Jul
        # tetap memengaruhi blok 30 Jun lewat aturan H+1.
        if refuse_h1_relevan(kind, status, paid, settle):
            report["refuse_mtoatd"].append({
                "cur": norm(currency), "completed": paid, "settle": settle,
                "usd": usd_existing, "tx": tx,
                "charges": (to_float(ws.cell(row=row, column=c_charges).value)
                            if c_charges else None),
                "source": source,
            })

        # --- SARINGAN BULAN LAPORAN ---------------------------------------
        # Basis tanggalnya SATU: kolom tanggal transaksi per sheet -- Paid Date
        # untuk D, Completed Date untuk W. Sama dengan yang dipakai seluruh
        # laporan, dan sama dengan cara tim Malaysia menyiapkan ekspornya.
        #
        # SUDAH DICOBA DAN DITOLAK (3 Sep 2026): ikut menyimpan baris yang
        # Settlement Date-nya di bulan ini walau tanggal transaksinya di luar,
        # dengan alasan baris MTOATD MT4入金/MT4出金/Wallet出金 dijumlah per
        # Settlement Date. Secara teori benar, tapi DIUJI terhadap sheet acuan
        # 'June 26 - MTOATD.xlsx' hasilnya JUSTRU RUSAK: sisi deposit yang tadinya
        # cocok 570/570 jadi 564/570, MT4出金 dari 570/570 jadi 562/570.
        # Sebabnya ekspor MEREKA juga disaring pakai tanggal transaksi, jadi
        # baris itu memang tidak ada di data mereka. Jangan diulang.
        if PERIODE_FILTER is not None and not dalam_periode(paid):
            kosongkan_kolom_hitung(ws, row, out)
            label = f"{paid:%Y-%m}" if paid else "(no date)"
            report["luar_periode"][(ws.title, label)] += 1
            report["luar_periode_usd"][(ws.title, label)] += usd_existing or 0.0
            n -= 1
            continue

        # Baris yang statusnya bukan 'finish' DIBUANG seluruhnya -- tidak masuk
        # laporan, pivot, Channel Balance, maupun total. Refuse H+1 yang relevan
        # sudah disalin khusus untuk MTOATD sebelum saringan periode/status.
        if HANYA_STATUS_FINISH and c_status and norm(status) not in STATUS_DIPAKAI:
            report["dibuang"][(ws.title, norm(currency), norm(status) or "(kosong)")] += 1
            report["dibuang_usd"][(ws.title, norm(currency), norm(status) or "(kosong)")] += \
                to_float(ws.cell(row=row, column=c_usd).value) or 0.0
            n -= 1
            continue
        ditolak = False

        # 1. Handling Fee
        komp, fixed, key, how = lookup_fee(fee_table, currency, gateway, kind)
        rate = komp["pct"] if komp else None
        fee = None
        if komp is not None and tx is not None:
            fee = tx * komp["pct"]
            if komp["min"] is not None:
                fee = max(fee, komp["min"])       # mis. '0.45%,min 90'
            fee += komp["fixed"]                  # mis. '1.5%+50'
            if INCLUDE_FIXED_FEE:
                fee += fixed                      # notasi lama '+10000' di kolom tambahan
        if ditolak:
            fee = None
            rate = None
            report["ditolak"][(norm(currency), norm(status))] += 1
        elif komp is None:
            report["fee_missing"][(norm(currency), norm(gateway))] += 1
        if komp is not None and not ditolak:
            report["fee_match"][(norm(currency), norm(gateway), f"{key[0]}/{key[1]}", how,
                                 komp["pct"], komp["fixed"], komp["min"], fixed, kind)] += 1

        # 2. Xero Rate
        xrate, used_date, fallback = lookup_rate(rates, xero_dates, paid, currency)
        if xrate is None:
            report["rate_missing"][(str(paid), norm(currency))] += 1
        elif fallback:
            report["rate_fallback"][(str(paid), norm(currency), str(used_date))] += 1
            kunci = (norm(currency), paid, used_date)
            det = report["fx_detail"].setdefault(
                kunci, {"rows": 0, "sheets": set(), "currency": norm(currency),
                        "missing_date": paid, "used_date": used_date})
            det["rows"] += 1
            det["sheets"].add(ws.title)

        # 3. Xero USD  &  4. Forex Gain/Loss
        xusd = None
        if xrate not in (None, 0) and tx is not None:
            xusd = tx * xrate if XERO_RATE_MULTIPLIER else tx / xrate
        forex = None
        if xusd is not None and usd_existing is not None:
            if kind == "withdrawal" and FLIP_FOREX_ON_WITHDRAWAL:
                forex = usd_existing - xusd
            else:
                forex = xusd - usd_existing

        stored_rate = None
        if xrate not in (None, 0):
            stored_rate = (1.0 / xrate) if XERO_RATE_MULTIPLIER else xrate

        for name, value, fmt in (
            ("Handling Fee", fee, FMT_FEE),
            ("Xero Rate", stored_rate, FMT_RATE),
            ("Xero USD", xusd, FMT_MONEY),
            ("Forex Gain/Loss", forex, FMT_MONEY),
        ):
            cell = ws.cell(row=row, column=out[name], value=value)
            cell.number_format = fmt
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font()
            if name == "Handling Fee" and ditolak:
                cell.value = None                    # ditolak -> tidak ada fee
                cell.fill = FILL_SKIP                # ABU-ABU = memang tidak kena fee
            elif name == "Handling Fee" and rate is None:
                cell.value = None                    # kosong, jangan diisi 0
                cell.fill = FILL_MISSING             # MERAH = rate belum ada
                cell.font = FONT_MISSING
            elif name in ("Xero Rate", "Xero USD", "Forex Gain/Loss") and value is None:
                cell.fill = FILL_MISSING             # MERAH = kurs tidak ketemu
                cell.font = FONT_MISSING
            elif name == "Xero Rate" and fallback:
                cell.fill = FILL_WARN                # KUNING = pakai kurs tanggal terdekat

        # keterangan per baris untuk kolom 'Data Status'
        catatan = []
        if ditolak:
            catatan.append(f"FEE: none - withdrawal {norm(status).lower()}, never paid out")
            # Kolom Remark biasanya berisi ALASAN penolakannya, mis.
            # 'Application failed. Please withdraw using THB currency.' -- itu
            # keterangan yang paling berguna buat orang yang membaca file hasil.
            c_remark = next((cols[k] for k in ("REMARK", "REMARKS") if k in cols), None)
            alasan = str(ws.cell(row=row, column=c_remark).value or "").strip() if c_remark else ""
            if alasan:
                catatan.append(f"REASON: {alasan}")
        elif rate is None:
            catatan.append(f"FEE: no rate for {norm(currency)} / {norm(gateway)}")
        if xrate is None:
            catatan.append(f"FX: no rate for {norm(currency)} on {paid}")
        elif fallback:
            selisih = (paid - used_date).days if (paid and used_date) else "?"
            catatan.append(f"FX: {paid} not in XERO, used {used_date} "
                           f"({selisih} day(s) earlier)")
        if tx is None:
            catatan.append("DATA: Transaction is empty")
        if not norm(gateway) and not ditolak:
            catatan.append("DATA: Payment Gateway is empty")
        if paid is None:
            catatan.append("DATA: date is empty or unreadable")

        # agregat untuk sheet 'D&W Report' -- currency & gateway APA ADANYA
        # (spasi di ujung tidak dibuang, supaya sama dgn perilaku pivot Excel)
        cur_raw = str(currency) if currency is not None else ""
        gw_raw = str(gateway) if gateway is not None else ""
        piv = report["pivot"].setdefault(
            (kind, cur_raw, gw_raw),
            {"crm_usd": 0.0, "tx": 0.0, "xero_usd": 0.0, "charges": 0.0, "n": 0,
             "urut": len(report["pivot"])})
        piv["n"] += 1
        if usd_existing is not None:
            piv["crm_usd"] += usd_existing
        if tx is not None:
            piv["tx"] += tx
        if xusd is not None:
            piv["xero_usd"] += xusd
        if c_charges:
            v_ch = to_float(ws.cell(row=row, column=c_charges).value)
            if v_ch:
                piv["charges"] += v_ch

        if MAKE_REPORT_SHEET:
            rec = {}
            for nama, kandidat in REPORT_COLUMNS:
                if kandidat is None:
                    rec[nama] = ws.cell(row=row, column=c_date).value
                else:
                    kol = next((cols[k] for k in kandidat if k in cols), None)
                    rec[nama] = ws.cell(row=row, column=kol).value if kol else None
            rec["Handling Fee"] = fee
            rec["Xero Rate"] = stored_rate
            rec["Xero USD"] = xusd
            rec["Forex Gain/Loss"] = forex
            rec["Source Sheet"] = ws.title
            rec["Date Source"] = date_col_name.title()
            rec["Type"] = kind
            rec["Data Status"] = "; ".join(catatan) if catatan else "OK"
            rec["_fee_missing"] = rate is None
            rec["_fx_missing"] = xrate is None
            rec["_fx_fallback"] = bool(fallback)
            report["records"].append(rec)

        if fee is not None:
            report["total_fee"][norm(currency)] += fee
            sub["fee"][norm(currency)] += fee
        if forex is not None:
            report["total_forex"] += forex
            sub["forex"] += forex
        if xusd is not None:
            report["total_xero_usd"] += xusd
            sub["xero_usd"] += xusd
        if usd_existing is not None:
            report["total_usd"] += usd_existing
            sub["usd"] += usd_existing
        sub["n"] += 1

    for name in NEW_COLUMNS:
        ws.column_dimensions[openpyxl.utils.get_column_letter(out[name])].width = max(14, len(name) + 3)
    if sys.stdout.isatty():
        print("\r".ljust(80), end="\r", flush=True)
    return n, date_col_name


def _header(ws, row, kolom, warna="4472C4"):
    from openpyxl.styles import Border, Side
    thin = Side(style="thin", color="D0D0D0")
    for c, t in enumerate(kolom, start=2):
        cell = ws.cell(row, c, t)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=warna)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def _rpt_urut(items):
    """Urutkan (currency, gateway) sesuai REPORT_GROUP_ORDER, currency dikelompokkan."""
    per_cur = {}
    for (cur, gw), v in items:
        per_cur.setdefault(cur, []).append((gw, v))

    if REPORT_GROUP_ORDER == "alpha":
        kunci_cur = lambda c: c
        kunci_gw = lambda t: t[0]
    elif REPORT_GROUP_ORDER == "appearance":
        kunci_cur = lambda c: min(v["urut"] for _, v in per_cur[c])
        kunci_gw = lambda t: t[1]["urut"]
    else:                                     # "value"
        kunci_cur = lambda c: -sum(v["crm_usd"] for _, v in per_cur[c])
        kunci_gw = lambda t: -t[1]["crm_usd"]

    hasil = []
    for cur in sorted(per_cur, key=kunci_cur):
        for gw, v in sorted(per_cur[cur], key=kunci_gw):
            hasil.append((cur, gw, v))
    return hasil


def write_fee_sheet(wb, fee_table):
    """Tulis hasil gabungan tabel fee jadi SATU sheet bernama FEE_SHEET.

    Layoutnya mengikuti sheet fee aslinya (header di baris 3) supaya familiar,
    ditambah kolom untuk biaya tetap / minimum dan kolom Source + Note yang
    memperlihatkan angka mana yang menimpa angka mana.

    Sheet fee LAIN dibuang setelah digabung, supaya file hasil cuma punya SATU
    tabel fee -- kalau tidak, file bisa berisi 'D&W FEE' (dari input) DAN
    'D&W TD FEE' (hasil gabungan) sekaligus, dan orang bingung mana yang dipakai."""
    from openpyxl.styles import Alignment

    for x in [x for x in wb.sheetnames
              if any(norm(x).startswith(a) for a in FEE_SHEET_AWALAN)]:
        del wb[x]
    ws = wb.create_sheet(FEE_SHEET)
    ws.sheet_properties.tabColor = "7030A0"

    ws["A1"] = "Dupoin Markets"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = " Handling Fee Rate"
    ws["A2"].font = Font(bold=True)

    kolom = ["Currency", "Payment Gateway", "Deposit", "Withdrawal",
             "Deposit Fixed", "Withdrawal Fixed", "Deposit Min", "Withdrawal Min",
             "Source", "Note"]
    for c, t in enumerate(kolom, start=1):
        sel = ws.cell(3, c, t)
        sel.font = Font(bold=True, color="FFFFFF")
        sel.fill = PatternFill("solid", fgColor="7030A0")
        sel.alignment = Alignment(wrap_text=True)

    r = 3
    for (cur, _k), v in sorted(fee_table.items(),
                               key=lambda x: (x[0][0], x[1].get("nama") or "")):
        r += 1
        dep, wd = v.get("deposit"), v.get("withdrawal")
        asal = v.get("asal") or []
        catat = list(v.get("ditimpa") or [])
        nilai = [
            cur, v.get("nama") or "",
            dep["pct"] if dep else None,
            wd["pct"] if wd else None,
            (dep["fixed"] or None) if dep else None,
            (wd["fixed"] or None) if wd else None,
            dep["min"] if dep else None,
            wd["min"] if wd else None,
            " + ".join(asal[-2:]) if len(asal) > 1 else (asal[0] if asal else ""),
            "; ".join(catat) or None,
        ]
        for c, val in enumerate(nilai, start=1):
            sel = ws.cell(r, c, val)
            if c in (3, 4):
                sel.number_format = "0.0000%"
            elif c in (5, 6, 7, 8):
                sel.number_format = "#,##0.##"
            if c == 10 and val:
                sel.fill = FILL_WARN
        info = next((v2 for (c2, g2), v2 in CATATAN_RATE.items()
                     if norm(c2) == cur and kunci_gw(g2) == _k), None)
        if info:
            sel = ws.cell(r, 10)
            sel.value = ((sel.value + "; ") if sel.value else "") + info
            sel.fill = PatternFill("solid", fgColor="E2EFDA")

        curiga = (fee_table.get((cur, _k)) or {}).get("curiga")
        if curiga:
            sel = ws.cell(r, 10)
            sel.value = ((sel.value + "; ") if sel.value else "") + \
                "LOOKS SWAPPED - PLEASE FIX IN YOUR FILE: " + curiga
            sel.fill = FILL_MISSING
            sel.font = FONT_MISSING

        perlu = next((v2 for (c2, g2), v2 in PERLU_KONFIRMASI.items()
                      if norm(c2) == cur and kunci_gw(g2) == _k), None)
        if perlu:
            sel = ws.cell(r, 10)
            sel.value = ((sel.value + "; ") if sel.value else "") + "NEEDS CONFIRMATION: " + perlu
            sel.fill = FILL_MISSING
            sel.font = FONT_MISSING
            for cc in (1, 2, 3, 4):
                ws.cell(r, cc).fill = FILL_WARN

        if v.get("fixed"):
            sel = ws.cell(r, 10)
            sel.value = ((sel.value + "; ") if sel.value else "") + \
                f"notasi lama '+{v['fixed']:g}' di sheet asal -- diterapkan hanya kalau " \
                f"INCLUDE_FIXED_FEE = True (sekarang {INCLUDE_FIXED_FEE})"
            sel.fill = FILL_WARN

    for kol, w in zip("ABCDEFGHIJ", (10, 22, 11, 12, 13, 15, 12, 14, 34, 62)):
        ws.column_dimensions[kol].width = w
    ws.freeze_panes = "A4"
    return r - 3


def nama_channel_bersih(gw, currency):
    """Nama channel untuk DITAMPILKAN di sheet 'Channel Balance': kode currency
    baris itu dibuang, di depan maupun di belakang.

        USDT / 'USDT-BEP20'    -> 'BEP20'
        NGN  / 'MNTX-NGN'      -> 'MNTX'
        PHP  / 'PA-PHP'        -> 'PA'
        USD  / 'Pay247-KH USD' -> 'Pay247-KH'

    Hanya kode currency BARIS ITU yang dibuang, bukan semua kode currency --
    'MNTX-LAK' di baris USD tetap utuh, karena LAK di situ memang berarti.
    Ini murni tampilan; pengelompokan tetap memakai kunci_channel()."""
    asli = str(gw or "").strip()
    kode = norm(currency)
    if not asli or not kode:
        return asli
    hasil = asli
    for sep in ("-", " ", ""):
        for pola in (sep + kode, kode + sep):
            if not pola:
                continue
            if pola.startswith(sep) and norm(hasil).endswith(norm(pola)):
                potong = hasil[: len(hasil) - len(pola)].strip(" -")
                if potong:
                    return potong
            if pola.endswith(sep) and norm(hasil).startswith(norm(pola)):
                potong = hasil[len(pola):].strip(" -")
                if potong:
                    return potong
    return hasil


# Nama channel yang SAMA tapi ditulis beda antara data transaksi dan sheet
# 'Payment Channel Balance' / 'Opening Balance'. Tanpa peta ini saldo pembukanya
# tidak ketemu dan saldo channel mulai dari NOL tanpa peringatan.
# Terbukti 26 Aug 2026: 6 dari 17 channel kehilangan saldo pembuka karena ini.
CHANNEL_ALIAS = {
    "MONETIX": "MNTX",          # data: 'MNTX-NGN'     sheet PCB: 'MONETIX'
    "THUNDERXPAY": "THX",       # data: 'THX-LAK'      sheet PCB: 'ThunderXpay'
    "THUNDERX": "THX",
}
# Alias yang hanya berlaku untuk satu currency (kunci: currency + nama apa adanya)
CHANNEL_ALIAS_CUR = {
    ("USD", "MNTX-BINANCE PAY"): "BINANCE PAY",     # sheet PCB menulis pakai spasi
    ("USD", "PAY247-KHR (USD)"): "PAY247-KH",       # data: 'PAY247-KH USD'
}

# Channel yang TIDAK ikut di sheet 'Channel Balance'. 'J Wallet' itu dompet USDT,
# bukan payment channel -- sheet 'Payment Channel Balance' mereka tidak punya kolom
# J Wallet sama sekali (channel USDT mereka: TRC20, BEP20, ERC20, DEFI PAY, MoPay,
# OTHERS). Pergerakan dompetnya sudah dilacak di sheet 'J Wallet (calc)'.
CHANNEL_BUKAN_PCB = ("J WALLET",)


def lewati_channel(currency, gateway):
    return any(x in norm(gateway) for x in CHANNEL_BUKAN_PCB)


def kunci_channel(currency, gateway):
    """Kunci channel yang sama untuk data transaksi maupun Fund Transfer Table.
    'VN77PAY' (data) dan '77PAY' (fund transfer) harus jadi kunci yang sama, jadi
    dipakai alias + buang suffix currency + buang tag （Manual）."""
    cur, gw = norm(currency), norm(gateway)
    # Tag （Manual） TIDAK dibuang di sini kecuali GABUNG_TAG_MANUAL = True:
    # sheet mereka mengeluarkan payout manual dari saldo channel sepenuhnya.
    # Untuk lookup rate fee tag itu SELALU dibuang -- lihat lookup_fee().
    if GABUNG_TAG_MANUAL:
        gw = strip_tags(gw)
    for cand in (CHANNEL_ALIAS_CUR.get((cur, gw), ""),
                 GATEWAY_ALIAS.get((cur, gw), ""),
                 strip_currency_suffix(gw, cur), gw):
        if cand:
            k = kunci_gw(cand)
            return CHANNEL_ALIAS.get(k, k)
    k = kunci_gw(gw)
    return CHANNEL_ALIAS.get(k, k)


def baca_fund_transfer(wb, rates=None, xero_dates=None, report=None):
    """-> list dict pemindahan dana dari sheet 'Fund Transfer Table'.

    Kalau PERIODE_FILTER diisi, baris di luar bulan itu DIBUANG dari hasil --
    tapi kolom rumus 'Xero (USD)' tetap dihitung ulang untuk SEMUA baris, karena
    kolom itu per-baris dan tidak ada hubungannya dengan bulan laporan.
    """
    nama = cari_sheet(wb, FTT_SHEET_NAMA)
    if nama is None:
        return None, None
    ws = wb[nama]
    baris = list(ws.iter_rows(max_row=min(ws.max_row, 20000), values_only=True))
    # Nama kolom BERULANG: sheet punya sisi 汇出 (dana keluar) dan 收款 (dana masuk)
    # dengan header yang sama (金额, 手续费, Xero (USD)). Peta dibuat PER SISI, dan
    # kunci duplikat mengambil kemunculan pertama di sisinya masing-masing.
    h_idx = kol = kol_in = None
    for i, row in enumerate(baris[:8]):
        nama_kol = [(norm(v), j) for j, v in enumerate(row) if v not in (None, "")]
        if not any(k.startswith("支付渠道") or k == "PAYMENT GATEWAY" for k, _ in nama_kol):
            continue
        batas = next((j for k, j in nama_kol if k.startswith("收款日期")), None)
        kol, kol_in = {}, {}
        for k, j in nama_kol:
            (kol_in if (batas is not None and j >= batas) else kol).setdefault(k, j)
        h_idx = i
        break
    if kol is None:
        return nama, []

    def _c(peta, *kandidat):
        for k in kandidat:
            for nm, j in peta.items():
                if nm == k or nm.startswith(k):
                    return j
        return None

    def c(*kandidat):
        return _c(kol, *kandidat)

    def ci(*kandidat):
        return _c(kol_in or {}, *kandidat)

    c_tgl = c("汇款日期", "DATE")
    c_dept = c("汇出部门", "DEPARTMENT")
    c_ch = c("支付渠道", "PAYMENT GATEWAY")
    c_cur = c("汇出币种", "CURRENCY")
    c_usd = c("XERO (USD)", "XERO USD")
    c_amt = c("金额", "AMOUNT")
    c_fee = c("手续费", "FEE")
    # sisi 收款 -> inilah yang menjadi baris J Wallet
    i_dept = ci("收款部门")
    # kolom penanda dompet -- isinya nama penerima, mis. J Wallet
    i_org = ci("收款人")
    i_cur = ci("收入币种", "CURRENCY")
    i_amt = ci("金额", "AMOUNT")

    out = []
    for row in baris[h_idx + 1:]:
        if c_ch is None or c_ch >= len(row):
            continue
        ch = norm(row[c_ch])
        d = as_date(row[c_tgl]) if c_tgl is not None and c_tgl < len(row) else None
        pen = (str(row[i_org]).strip()
               if i_org is not None and i_org < len(row) and row[i_org] else "")
        # Baris yang sisi 汇出-nya KOSONG tapi sisi 收款-nya ada tetap dipakai:
        # itu dana MASUK ke sebuah channel tanpa channel pengirim (mis. 1 Jun 2026
        # 'Client Test' -> Kotak Bank2 sebesar 1). Dulu baris seperti ini dibuang,
        # jadi netto Fund transfer Kotak Bank2 salah 1 (282.251 bukan 282.250).
        if not d or (not ch and not pen):
            continue
        out.append({
            "tgl": d,
            "dept": row[c_dept] if c_dept is not None and c_dept < len(row) else None,
            "channel": str(row[c_ch]).strip() if ch else None,
            "currency": norm(row[c_cur]) if c_cur is not None and c_cur < len(row) else "",
            "usd": to_float(row[c_usd]) if c_usd is not None and c_usd < len(row) else None,
            "_baris": h_idx + 1 + len(out) + 1,   # nomor baris di sheet, utk tulis balik
            "jumlah": to_float(row[c_amt]) if c_amt is not None and c_amt < len(row) else None,
            "fee": to_float(row[c_fee]) if c_fee is not None and c_fee < len(row) else None,
            "tujuan": (str(row[i_dept]).strip()
                       if i_dept is not None and i_dept < len(row) and row[i_dept] else None),
            "penerima": (str(row[i_org]).strip()
                         if i_org is not None and i_org < len(row) and row[i_org] else None),
            "tujuan_cur": (norm(row[i_cur])
                           if i_cur is not None and i_cur < len(row) else ""),
            "tujuan_jumlah": (to_float(row[i_amt])
                              if i_amt is not None and i_amt < len(row) else None),
        })
    # Kolom 'Xero (USD)' sisi 汇出 BUKAN data -- itu satu-satunya kolom berumus di
    # 'Fund Transfer Table':
    #     Xero (USD) = 金额 / kurs Xero pada tanggal itu untuk currency itu
    # Sama dengan kolom 'Xero USD' di sheet D/W. Karena itu KOLOM RUMUS, isinya
    # DIHITUNG ULANG setiap kali dijalankan -- bukan cuma diisi kalau kosong.
    # Nilai lama yang berbeda ditimpa (mis. kurs sudah dikoreksi di sheet Xero, atau
    # sel warisan berisi #VALUE!). Kalau kursnya tidak ketemu, nilai lama DIPERTAHANKAN
    # supaya tidak ada angka yang hilang, dan jumlahnya dilaporkan.
    n_isi = n_koreksi = n_tanpa_kurs = n_dikosongkan = 0
    if rates is not None and xero_dates:
        for t in out:
            if not t["jumlah"] or not t["currency"] or not t["tgl"]:
                # Tidak ada 金额 / currency / tanggal di sisi 汇出 -> tidak ada yang
                # bisa dihitung, jadi selnya harus KOSONG. Angka sisa dari template
                # yang salah tidak boleh dibiarkan menetap di kolom rumus.
                if t["usd"] is not None:
                    t["usd"] = None
                    n_dikosongkan += 1
                    if c_usd is not None:
                        # JEBAKAN openpyxl: ws.cell(r, c, value=None) TIDAK menghapus
                        # isi sel -- None diartikan "tidak ada nilai yang diberikan",
                        # jadi nilai lamanya tetap. Harus lewat .value.
                        ws.cell(row=t["_baris"], column=c_usd + 1).value = None
                continue
            kurs, _dipakai, _fb = lookup_rate(rates, xero_dates, t["tgl"], t["currency"])
            if kurs in (None, 0):
                if t["usd"] is None:
                    n_tanpa_kurs += 1
                continue
            baru = t["jumlah"] / kurs
            lama = t["usd"]
            if lama is None:
                n_isi += 1
            elif abs(lama - baru) > 0.005:
                n_koreksi += 1
            t["usd"] = baru
            if c_usd is not None:                       # tulis balik ke sheet
                sel = ws.cell(row=t["_baris"], column=c_usd + 1, value=baru)
                sel.number_format = FMT_MONEY
    if n_isi or n_koreksi or n_tanpa_kurs or n_dikosongkan:
        bagian = []
        if n_isi:
            bagian.append(f"{n_isi} kosong diisi")
        if n_koreksi:
            bagian.append(f"{n_koreksi} nilai lama dikoreksi")
        if n_dikosongkan:
            bagian.append(f"{n_dikosongkan} dikosongkan (tidak ada jumlah/currency)")
        if n_tanpa_kurs:
            bagian.append(f"{n_tanpa_kurs} tidak bisa (kurs tidak ada di sheet Xero)")
        print(f"         'Xero (USD)' Fund Transfer Table dihitung ulang "
              f"(jumlah / kurs Xero): {', '.join(bagian)}")

    # Saringan bulan laporan. Dilakukan SETELAH 'Xero (USD)' dihitung ulang di
    # atas, supaya kolom rumus itu tetap benar di seluruh sheet, bukan cuma di
    # baris bulan ini. Yang tersaring hanya AGREGATNYA: J Wallet dan kolom
    # 'Fund transfer' di Payment Channel Balance.
    if PERIODE_FILTER is not None:
        semua = len(out)
        dibuang = Counter()
        simpan = []
        for t in out:
            if dalam_periode(t["tgl"]):
                simpan.append(t)
            else:
                dibuang[f"{t['tgl']:%Y-%m}" if t["tgl"] else "(no date)"] += 1
        out = simpan
        if dibuang:
            rinci = ", ".join(f"{k} {v:,}" for k, v in sorted(dibuang.items()))
            print(f"         Fund Transfer Table: {semua - len(out):,} dari {semua:,} baris "
                  f"di luar {nama_periode()} dilewati ({rinci})")
            if report is not None:
                for k, v in dibuang.items():
                    report["luar_periode"][("Fund Transfer Table", k)] += v
    return nama, out


OB_SHEET = "Opening Balance"          # dibuat oleh buat_template / isi_template

# Dari sheet mana saldo pembuka benar-benar terbaca. Diisi oleh saldo_awal_*
# supaya pesan di Terminal menyebut sumber yang BENAR -- sheet aslinya bisa ada
# tapi tidak memuat baris sebelum tanggal data, lalu jatuh ke 'Opening Balance'.
_ASAL_SALDO = {"jw": None, "channel": None}
OB_JWALLET_CHANNEL = "J WALLET"       # penanda baris dompet USDT di sheet itu


def baca_opening_balance(wb):
    """Sheet 'Opening Balance' template -> ({(cur, channel): saldo}, tgl, saldo_jw, tgl_jw).

    Dipakai kalau workbook yang dihitung TIDAK memuat sheet 'J Wallet' /
    'Payment Channel Balance' aslinya -- itu kasus normal untuk template."""
    nama = next((x for x in wb.sheetnames if norm(x) == norm(OB_SHEET)), None)
    if nama is None:
        return {}, None, None, None
    ws = wb[nama]
    # Header harus dicocokkan TEPAT. Kalau pakai startswith, baris catatan di
    # atasnya ("Opening balances actually used for this run...") ikut dikira header
    # dan seluruh sheet terbaca kosong.
    kol = None
    for i, row in enumerate(ws.iter_rows(max_row=6, values_only=True)):
        nama_kol = {norm(v): j for j, v in enumerate(row) if v}
        if "OPENING BALANCE" in nama_kol and "CURRENCY" in nama_kol:
            kol = nama_kol
            baris_awal = i + 2
            break
    if not kol:
        return {}, None, None, None
    c_tgl = kol.get("AS OF DATE")
    c_cur, c_ch = kol.get("CURRENCY"), kol.get("CHANNEL")
    c_bal = kol.get("OPENING BALANCE")
    if c_cur is None or c_ch is None or c_bal is None:
        return {}, None, None, None

    out, tgl, jw, tgl_jw = {}, None, None, None
    for row in ws.iter_rows(min_row=baris_awal, values_only=True):
        if c_bal >= len(row) or not isinstance(row[c_bal], (int, float)):
            continue
        cur = norm(row[c_cur]) if c_cur < len(row) else ""
        ch = str(row[c_ch] or "").strip() if c_ch < len(row) else ""
        d = as_date(row[c_tgl]) if c_tgl is not None and c_tgl < len(row) else None
        if norm(ch) == OB_JWALLET_CHANNEL:
            jw, tgl_jw = float(row[c_bal]), d
            continue
        if not cur or not ch:
            continue
        out[(cur, kunci_channel(cur, ch))] = float(row[c_bal])
        tgl = d or tgl
    return out, tgl, jw, tgl_jw


def saldo_awal_jwallet(wb, tgl_awal):
    """Saldo pembuka dompet USDT: baris 余额 terakhir SEBELUM tgl_awal di sheet
    'J Wallet' yang sudah ada di workbook. Kalau sheet itu tidak ada, dibaca dari
    sheet 'Opening Balance'. -> (saldo, tanggal) atau (0.0, None)."""
    # Angka yang diketik manusia MENANG atas semua sumber lain -- itu memang
    # gunanya: dipakai justru ketika file yang diupload tidak punya sheet saldo.
    if SALDO_AWAL_JW is not None:
        _ASAL_SALDO["jw"] = "opening balance you typed in"
        tgl = tanggal_saldo_pembuka()
        if tgl is None and tgl_awal:
            tgl = tgl_awal - datetime.timedelta(days=1)
        return SALDO_AWAL_JW, tgl

    def _dari_ob():
        _c, _t, jw, tgl_jw = baca_opening_balance(wb)
        _ASAL_SALDO["jw"] = OB_SHEET if jw is not None else None
        return (jw, tgl_jw) if jw is not None else (0.0, None)

    nama = next((x for x in wb.sheetnames if norm(x) == "J WALLET"), None)
    if nama is None or tgl_awal is None:
        return _dari_ob()
    ws = wb[nama]
    # Kolom saldo dicari dari NAMA header (余额 / balance). Jangan pakai
    # "angka terakhir dari kanan": banyak baris kolom saldonya kosong sementara
    # kolom 转出 di sebelahnya berisi angka -> saldo pembuka jadi salah.
    h_idx = c_bal = None
    for i, row in enumerate(ws.iter_rows(max_row=8, values_only=True)):
        if not any(norm(v) == "CURRENCY" for v in row if v):
            continue
        h_idx = i
        for j, v in enumerate(row):
            if v and ("余额" in str(v) or norm(v).startswith("BALANCE")):
                c_bal = j
        break
    if h_idx is None or c_bal is None:
        return 0.0, None
    hasil = (0.0, None)
    for row in ws.iter_rows(min_row=h_idx + 2, values_only=True):
        d = as_date(row[0]) if row else None
        if not d or d >= tgl_awal or c_bal >= len(row):
            continue
        if isinstance(row[c_bal], (int, float)):
            hasil = (float(row[c_bal]), d)
    # Sheet 'J Wallet' ada tapi tidak memuat baris SEBELUM tanggal awal -- itu yang
    # terjadi kalau yang diupload adalah FILE HASIL kita sendiri (sheet 'J Wallet'
    # di situ mulai dari tanggal data, bukan sebelumnya). Jatuh ke 'Opening Balance'.
    if hasil[1] is None:
        return _dari_ob()
    _ASAL_SALDO["jw"] = nama
    return hasil


def saldo_awal_channel(wb, tgl_awal):
    """Saldo pembuka per (currency, channel) dari blok 余额 sheet
    'Payment Channel Balance' -- baris terakhir sebelum tgl_awal."""
    def _dari_ob():
        saldo, tgl, _jw, _tjw = baca_opening_balance(wb)
        _ASAL_SALDO["channel"] = OB_SHEET if saldo else None
        return saldo, tgl

    nama = next((x for x in wb.sheetnames
                 if norm(x).startswith("PAYMENT CHANNEL BALANCE")), None)
    if nama is None or tgl_awal is None:
        return _dari_ob()
    ws = wb[nama]
    baris = list(ws.iter_rows(values_only=True))
    if len(baris) < 6:
        return _dari_ob()
    r4, r5 = baris[3], baris[4]
    # blok 余额: cari label 余额 di baris 2-3 (sel gabungan -> nilai di kiri-atas)
    awal = None
    for r in (baris[1], baris[2]):
        for j, v in enumerate(r or []):
            if v and "余额" in str(v):
                awal = j
    if awal is None:
        return _dari_ob()
    cur = None
    kol = {}
    for j in range(awal, len(r5)):
        if j < len(r4) and r4[j] not in (None, ""):
            cur = norm(r4[j])
        if r5[j] not in (None, "") and cur:
            kol[(cur, kunci_channel(cur, str(r5[j])))] = j

    out, tgl = {}, None
    for row in baris[5:]:
        d = as_date(row[0]) if row else None
        if not d or d >= tgl_awal:
            continue
        tgl = d
        for k, j in kol.items():
            if j < len(row) and isinstance(row[j], (int, float)):
                out[k] = float(row[j])
    if not out:                      # layout tidak dikenali / tidak ada baris sebelumnya
        return _dari_ob()            # (mis. yang diupload file HASIL kita sendiri)
    _ASAL_SALDO["channel"] = nama
    return out, tgl


def baris_pesan_jwallet_kosong(rr, n_trx):
    return rr + 1 if n_trx == 0 else None


def write_channel_sheets(wb, report, ftt):
    """Sheet 'Channel Balance' + 'J Wallet (calc)'.

    Ketujuh sub-blok dihitung. 'Withdrawal charges' ditandai KUNING bukan karena
    ragu -- rumusnya sudah dikonfirmasi tim 3 Sep 2026 -- tapi karena kolom itu
    boleh disesuaikan manual oleh tim keuangan setelah laporan jadi."""
    from openpyxl.styles import Alignment

    # --- agregat D/W per (tanggal, currency, channel)
    # Payment Channel Balance memakai Paid Date untuk deposit dan Apply Date
    # untuk withdrawal (terbukti cocok persis di 1 Jun 2026) -- BUKAN Completed Date.
    agg = defaultdict(lambda: defaultdict(float))
    for rec in report["records"]:
        cur, gw = norm(rec.get("Currency")), str(rec.get("Payment Gateway") or "").strip()
        if not cur or not gw:
            continue
        # Basis tanggal: deposit pakai Paid Date, withdrawal pakai Completed Date
        # (dikonfirmasi tim Malaysia). rec["Date"] sudah kolom tanggal yang dipakai
        # per sheet -- Paid Date untuk D, Completed Date untuk W.
        d = as_date(rec.get("Date"))
        if not d:
            continue
        tx = to_float(rec.get("Transaction")) or 0.0
        fee = to_float(rec.get("Handling Fee"))
        k = (d, cur, kunci_channel(cur, gw))
        agg[k]["_nama"] = agg[k].get("_nama") or gw
        if rec["Type"] == "withdrawal":
            agg[k]["wd"] += tx
            if fee:
                agg[k]["wd_fee"] += fee
        else:
            agg[k]["dep"] += tx
            if fee:
                agg[k]["dep_fee"] += fee

    # --- pemindahan dana: NETTO (keluar dikurangi masuk) per channel
    # Sheet 'Payment Channel Balance' mereka memakai netto, bukan cuma sisi keluar.
    # Terbukti 1 Jun 2026:
    #   INR / Kotak Bank2 : keluar 394.100, masuk 111.849 + 1 -> mereka tulis 282.250
    #   USDT / TRC20      : keluar 0, masuk 30.000            -> mereka tulis -30.000
    # Kolom ini DIKURANGKAN dari saldo, jadi tanda positif = dana keluar channel.
    for t in (ftt or []):
        if t.get("channel") and not lewati_channel(t["currency"], t["channel"]):
            k = (t["tgl"], t["currency"], kunci_channel(t["currency"], t["channel"]))
            agg[k]["_nama"] = agg[k].get("_nama") or t["channel"]
            agg[k]["ft"] += t["jumlah"] or 0.0
            agg[k]["ft_fee"] += t["fee"] or 0.0
        # sisi penerima: dana MASUK ke channel itu -> netto berkurang
        ch_in = t.get("penerima")
        cur_in = t.get("tujuan_cur") or t["currency"]
        jml_in = t.get("tujuan_jumlah")
        if ch_in and not lewati_channel(cur_in, ch_in) and jml_in:
            k = (t["tgl"], cur_in, kunci_channel(cur_in, ch_in))
            agg[k]["_nama"] = agg[k].get("_nama") or ch_in
            agg[k]["ft"] -= jml_in

    if not agg:
        return 0, 0

    # SALDO PEMBUKA DIBACA DULU. Nama sheet hasil kita sama dengan sheet input
    # ('Payment Channel Balance', 'J Wallet'), jadi kalau sheet lamanya dihapus
    # duluan, sumber saldo pembukanya ikut hilang dan saldo mulai dari NOL.
    tgl_awal = min((d for d, _c, _g in agg), default=None)
    awal_ch, tgl_ch = saldo_awal_channel(wb, tgl_awal)
    asal_ch = _ASAL_SALDO["channel"] or OB_SHEET
    tgl_jw = min((t["tgl"] for t in (ftt or [])), default=None)
    bal, tgl_bal = saldo_awal_jwallet(wb, tgl_jw)
    asal_jw = _ASAL_SALDO["jw"] or OB_SHEET

    # ---------------- Payment Channel Balance ----------------
    if PCB_SHEET in wb.sheetnames:
        del wb[PCB_SHEET]
    ws = wb.create_sheet(PCB_SHEET)
    ws.sheet_properties.tabColor = "C55A11"
    for c, t in enumerate(PCB_KOLOM, start=1):
        sel = ws.cell(1, c, t)
        sel.font = Font(bold=True, color="FFFFFF", size=10)
        sel.fill = PatternFill("solid", fgColor="C55A11")
        sel.alignment = Alignment(wrap_text=True, horizontal="center")
    sel_wc = ws.cell(1, PCB_KOLOM.index("Withdrawal charges") + 1)
    sel_wc.fill = PatternFill("solid", fgColor="BF8F00")
    sel_wc.comment = Comment(INFO_WITHDRAWAL_CHARGES, "hitung_dw.py", height=300, width=460)
    ws.row_dimensions[1].height = 30

    kuning = PatternFill("solid", fgColor="FFF2CC")
    saldo = defaultdict(float)
    saldo.update(awal_ch)
    if not awal_ch:
        print(f"Lewati : saldo pembuka channel tidak ketemu -> kolom 'Balance' di "
              f"'{PCB_SHEET}' mulai dari NOL")
        report["input_kurang"].append((
            "Opening Balance",
            "no opening balances found (sheet 'Opening Balance' is empty, and the "
            "workbook has no 'Payment Channel Balance' / 'J Wallet' sheet to read them from)",
            f"The 'Balance' column of '{PCB_SHEET}' and '{JW_SHEET}' starts from ZERO, "
            f"so the running balances are movement-only, not real balances.",
            "Fill the 'Opening Balance' sheet of the template (last closing balance "
            "before this month, per currency + channel, plus one row with Channel = "
            "'J Wallet' for the USDT wallet), or run isi_template.py which fills it "
            "automatically from the source workbook."))
    if awal_ch:
        print(f"         saldo pembuka channel: {len(awal_ch)} kombinasi dari "
              f"sheet {asal_ch!r} per {tgl_ch}")
    # Nama tampilan: kode currency dibuang. Kalau pembersihan itu membuat dua channel
    # berbeda jadi bernama sama, dua-duanya dikembalikan ke nama aslinya supaya
    # tidak ada dua baris yang terlihat identik.
    nama_bersih, bentrok = {}, defaultdict(set)
    for (d, cur, kgw), v in agg.items():
        asli = v.get("_nama", kgw)
        nama_bersih[(cur, kgw)] = nama_channel_bersih(asli, cur)
        bentrok[(cur, nama_channel_bersih(asli, cur))].add(kgw)
    for (cur, nm), kunci in bentrok.items():
        if len(kunci) > 1:
            for kgw in kunci:
                asli = next(v.get("_nama", kgw) for (d2, c2, k2), v in agg.items()
                            if c2 == cur and k2 == kgw)
                nama_bersih[(cur, kgw)] = asli

    r = 1
    n_kuning = 0
    for (d, cur, kgw) in sorted(agg, key=lambda x: (x[0], x[1], x[2])):
        v = agg[(d, cur, kgw)]
        r += 1
        dep, dfee = v.get("dep", 0.0), v.get("dep_fee", 0.0)
        ft, ffee = v.get("ft", 0.0), v.get("ft_fee", 0.0)
        wd = v.get("wd", 0.0)
        # Withdrawal charges = jumlah Handling Fee baris withdrawal (Transaction x rate
        # withdrawal dari tabel fee). Terbukti PERSIS untuk banyak channel:
        # LAK/ThunderXpay 1,0000% di 132 hari, INR/MONETIX 4,5% di 36 hari,
        # UZS/MONETIX 3,5% di 26 hari, INR/BeckPay 2,0% di 10 hari.
        # Beberapa channel di sheet MEREKA angkanya lebih besar, TAPI itu bukan
        # kelebihan yang salah: sel-sel itu memang DIKETIK MANUAL, tidak ada
        # rumusnya. Dikonfirmasi tim 3 Sep 2026: "continue to use the same formula,
        # our side will check with other reports and edit those that needs to be
        # edited manually". Jadi rumus ini yang benar; KUNING sekarang cuma penanda
        # bahwa kolom ini boleh disesuaikan tangan setelah laporan jadi.
        wfee = v.get("wd_fee", 0.0) or None
        kunci_saldo = (cur, kgw)
        saldo[kunci_saldo] += dep - dfee - ft - ffee - wd - (wfee or 0.0)
        catatan = ("Withdrawal charges = withdrawal amount x withdrawal rate from the fee "
                   "table. Confirmed 3 Sep 2026 as the formula to use; the finance team "
                   "adjusts individual channels by hand afterwards where a gateway "
                   "statement differs.")
        if (cur, kgw) not in awal_ch:
            catatan += (" No opening balance found for this channel, so Balance starts "
                        "from zero.")
        nilai = [d, cur, nama_bersih.get((cur, kgw), v.get("_nama", kgw)),
                 dep, dfee, ft, ffee, wd, wfee,
                 saldo[kunci_saldo], catatan]
        for c, val in enumerate(nilai, start=1):
            sel = ws.cell(r, c, val)
            if c == 1:
                sel.number_format = "yyyy-mm-dd"
            elif c >= 4 and c <= 10:
                sel.number_format = FMT_ACC
            if c == PCB_KOLOM.index("Withdrawal charges") + 1:
                sel.fill = kuning
                n_kuning += 1
            elif c == PCB_KOLOM.index("Balance") + 1:
                sel.fill = PatternFill("solid", fgColor="E2EFDA")
    for kol, w in zip("ABCDEFGHIJK", (12, 10, 22, 16, 16, 16, 18, 16, 18, 18, 48)):
        ws.column_dimensions[kol].width = w
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:K{max(r, 2)}"
    n_pcb = r - 1

    # ---------------- J Wallet ----------------
    if JW_SHEET in wb.sheetnames:
        del wb[JW_SHEET]
    wj = wb.create_sheet(JW_SHEET)
    wj.sheet_properties.tabColor = "7030A0"
    KOL = ["Date", "TD", "Currency", "IN/OUT", "Transfer out (original)",
           "Transfer in (USD)", "Transfer out (USD)", "Balance (USD)", "Source"]
    for c, t in enumerate(KOL, start=1):
        sel = wj.cell(1, c, t)
        sel.font = Font(bold=True, color="FFFFFF", size=10)
        sel.fill = PatternFill("solid", fgColor="7030A0")
        sel.alignment = Alignment(wrap_text=True, horizontal="center")
    wj.row_dimensions[1].height = 30

    # Baris mana yang menggerakkan dompet, dan BERAPA. Rumus mereka (dari komentar
    # Liaw Zern di sheet 'J Wallet') selalu mengambil SISI USDT:
    #   转入 = SUMIFS(FTT!M:M, 收入币种="USDT", 收款人="J Wallet", ...)   <- 金额 MASUK
    #   转出 = -SUMIFS(FTT!G:G, 汇出币种="USDT", 支付渠道="*J Wallet*", ...) <- 金额 KELUAR
    # Jadi:
    #   dana MASUK dompet -> pakai 金额 sisi 收款   (harus USDT)
    #   dana KELUAR dompet -> pakai 金额 sisi 汇出  (harus USDT)
    # JEBAKAN: satu baris bisa USDT di satu sisi dan mata uang lain di sisi lain.
    # Contoh nyata 14 Feb 2026: J Wallet USDT 30.000 -> 77PAY VND 792.000.000.
    # Kalau yang diambil sisi 收款, saldo dompet USDT kemasukan angka VND dan
    # hasilnya ngawur (dulu saldo tutup jadi -4,75 miliar). Data 1 hari tidak
    # memperlihatkan ini karena kebetulan kedua sisinya 30.000 USDT.
    USDT = "USDT"

    def sisi_wallet(t):
        """-> (arah, jumlah USDT) atau None kalau baris ini tidak menyentuh dompet."""
        masuk_ke = ("J WALLET" in norm(t.get("penerima"))
                    or "J WALLET" in norm(t.get("tujuan")))
        keluar_dari = "J WALLET" in norm(t["channel"])
        if masuk_ke and not keluar_dari:
            if norm(t.get("tujuan_cur")) == USDT:
                return "IN", (t.get("tujuan_jumlah") or 0.0)
            return None                      # masuk dompet tapi bukan USDT -> bukan dompet USDT
        if keluar_dari:
            if norm(t.get("currency")) == USDT:
                return "OUT", (t.get("jumlah") or 0.0)
            return None
        return None

    def sentuh_wallet(t):
        return sisi_wallet(t) is not None

    if tgl_bal:
        asal_txt = ("diketik di halaman upload" if SALDO_AWAL_JW is not None
                    else f"dari sheet {asal_jw!r}")
        print(f"         saldo pembuka J Wallet: {bal:,.2f} ({asal_txt} per {tgl_bal})")
    bal_awal = bal          # 'bal' berubah di dalam loop -> simpan yang pembuka
    rr = 1
    n_trx = 0               # baris TRANSAKSI saja, tidak termasuk baris saldo pembuka

    # BARIS PERTAMA = saldo penutup bulan sebelumnya, supaya kelihatan dari mana
    # angka saldo berjalan ini dimulai. Tanpa baris ini, saldo pembuka cuma
    # 'tersembunyi' di dalam angka Balance baris transaksi pertama dan orang yang
    # membaca sheet tidak bisa melihat titik mulainya.
    if bal_awal or tgl_bal:
        rr += 1
        asal_txt = ("opening balance entered on the upload page"
                    if SALDO_AWAL_JW is not None else f"from sheet '{asal_jw}'")
        pembuka = [tgl_bal, "Opening balance", "USDT", None, None, None, None,
                   bal_awal, asal_txt]
        for c, val in enumerate(pembuka, start=1):
            sel = wj.cell(rr, c, val)
            sel.font = Font(bold=True)
            sel.fill = PatternFill("solid", fgColor="FFF2CC")   # kuning muda
            if c == 1:
                sel.number_format = "yyyy-mm-dd"
            elif c >= 5:
                sel.number_format = FMT_ACC
        wj.cell(rr, 9).font = Font(bold=True, italic=True, size=9, color="808080")

    # JEBAKAN: 'channel' (sisi 汇出) BOLEH None -- baris dana MASUK tanpa channel
    # pengirim memang sengaja dipertahankan (lihat baca_fund_transfer). Kalau dua
    # baris di tanggal yang sama punya channel None dan string, sorted() mencoba
    # membandingkan None < str dan CRASH. Terjadi nyata di data Juli 2026.
    for t in sorted((x for x in (ftt or []) if sentuh_wallet(x)),
                    key=lambda x: (x["tgl"], x["channel"] or "")):
        rr += 1
        n_trx += 1
        # Jumlah USDT yang benar-benar masuk/keluar dompet = sisi 收款 (kolom 金额
        # di paruh kanan sheet). Kurs 'Xero (USD)' hanya nilai referensi, biasanya
        # sedikit beda (55,162.35 vs 55,000 yang benar-benar diterima).
        arah, jml = sisi_wallet(t)
        # Kolom 'TD' = PIHAK LAWAN dompet, bukan dompetnya sendiri -- sama seperti
        # sheet 'J Wallet' mereka. Kalau dana KELUAR dari dompet, 支付渠道 isinya
        # 'J Wallet', jadi yang harus ditampilkan adalah penerimanya (收款人),
        # mis. 'TRC20'. Kalau dana MASUK, penerimanya 'J Wallet', jadi yang
        # ditampilkan channel pengirimnya (mis. 'OCPAY').
        if arah == "OUT":
            lawan = (t.get("penerima") or t.get("tujuan") or t["channel"])
        else:
            # channel pengirim boleh kosong -- pakai departemen pengirim sebagai
            # label supaya kolom 'TD' tidak melompong (mis. 'Client Test').
            lawan = t["channel"] or t.get("dept")
        if "J WALLET" in norm(lawan):        # jaga-jaga: jangan tampilkan dompetnya sendiri
            lawan = next((x for x in (t.get("penerima"), t.get("tujuan"), t["channel"])
                          if x and "J WALLET" not in norm(x)), lawan)
        masuk = jml if arah == "IN" else 0.0
        keluar = -jml if arah == "OUT" else 0.0
        bal += masuk + keluar

        # Kolom 'Currency' dan 'Transfer out (original)' selalu dari sisi PIHAK
        # LAWAN, mengikuti rumus mereka:
        #   E = IF(IN, -SUMIFS(FTT!G:G,...),  <- IN : minus 金额 sisi 汇出
        #              SUMIFS(FTT!M:M,...))   <- OUT: plus  金额 sisi 收款
        # Jadi tandanya: IN negatif, OUT POSITIF. Dan mata uangnya ikut lawan --
        # untuk baris OUT itu 收入币种, bukan 汇出币种. Contoh 14 Feb 2026:
        # J Wallet USDT 30.000 -> 77PAY VND 792.000.000 ditampilkan sebagai
        # "77PAY | VND | OUT | 792.000.000", bukan "USDT | -30.000".
        if arah == "OUT":
            cur_tampil = t.get("tujuan_cur") or t["currency"]
            asli = t.get("tujuan_jumlah")
            if asli is None:
                asli = t.get("jumlah")
            asli = abs(asli) if asli else None
        else:
            cur_tampil = t["currency"]
            asli = -(t["jumlah"] or 0.0) if t["jumlah"] else None

        nilai = [t["tgl"], lawan, cur_tampil, arah, asli,
                 masuk or None, keluar or None, bal, t["dept"]]
        for c, val in enumerate(nilai, start=1):
            sel = wj.cell(rr, c, val)
            if c == 1:
                sel.number_format = "yyyy-mm-dd"
            elif c >= 5:
                sel.number_format = FMT_ACC
            if c == 8:
                sel.fill = PatternFill("solid", fgColor="E2EFDA")
    if n_trx == 0:
        # Pesan sesudah opening balance; jangan menimpa baris pembuka.
        msg_row = baris_pesan_jwallet_kosong(rr, n_trx)
        wj.cell(msg_row, 1, "EMPTY - no J Wallet transactions found.").font = \
            Font(bold=True, color="9C0006")
        wj.cell(msg_row + 1, 1, "This sheet only includes USDT transfers to or from J Wallet "
                      "found in 'Fund Transfer Table'. No matching rows were found. That data "
                      "cannot be derived from sheets D or W. See 'Missing Data', section 3.").font = \
            Font(italic=True, size=10, color="808080")
    for kol, w in zip("ABCDEFGHI", (12, 22, 10, 9, 24, 18, 18, 18, 28)):
        wj.column_dimensions[kol].width = w
    wj.freeze_panes = "C2"

    # Tulis saldo pembuka YANG DIPAKAI ke sheet 'Opening Balance' di file hasil.
    # Tanpa ini, file hasil yang diproses ulang kehilangan saldo pembukanya --
    # sheet 'J Wallet' / 'Payment Channel Balance' hasil kita mulai dari tanggal
    # data, jadi tidak ada baris "sebelum" yang bisa dibaca.
    if awal_ch or bal_awal:
        if OB_SHEET in wb.sheetnames:
            del wb[OB_SHEET]
        wo = wb.create_sheet(OB_SHEET)
        wo.sheet_properties.tabColor = "7F6000"
        wo.cell(1, 1, "Opening balances actually used for this run. Written automatically "
                      "so that re-running this result file keeps the same starting point.")\
            .font = Font(italic=True, size=9, color="808080")
        for c, t in enumerate(["As Of Date", "Currency", "Channel", "Opening Balance"], 1):
            sel = wo.cell(2, c, t)
            sel.font = Font(bold=True, color="FFFFFF", size=10)
            sel.fill = PatternFill("solid", fgColor="7F6000")
        r_ob = 2
        if bal_awal:
            r_ob += 1
            for c, v in enumerate((tgl_bal, "USDT", "J Wallet", bal_awal), 1):
                sel = wo.cell(r_ob, c, v)
                if c == 1:
                    sel.number_format = "yyyy-mm-dd"
                elif c == 4:
                    sel.number_format = FMT_ACC
        for (cur_o, ch_o), v in sorted(awal_ch.items()):
            r_ob += 1
            for c, x in enumerate((tgl_ch, cur_o, ch_o, v), 1):
                sel = wo.cell(r_ob, c, x)
                if c == 1:
                    sel.number_format = "yyyy-mm-dd"
                elif c == 4:
                    sel.number_format = FMT_ACC
        for kol, w in zip("ABCD", (12, 10, 22, 18)):
            wo.column_dimensions[kol].width = w
        wo.freeze_panes = "A3"

    return n_pcb, n_trx


def write_mtoatd_sheet(wb, report):
    """Sheet 'MTOATD': rekonsiliasi harian, dihitung PENUH dari sheet D + W.

    Tidak ada lagi tahap input manual. Semua rumusnya ada di mtoatd_spec.py --
    dibongkar dari sheet 'MTOATD (MAY''26)' di workbook sumber dan terverifikasi
    523/523 sel. Baris yang di workbook sumber tidak punya rumus (rincian manual
    seperti ——作废单, ——小数点差异, Wallet入金) dibiarkan KOSONG."""
    from openpyxl.styles import Alignment
    import mtoatd_spec as spec

    if MTOATD_SHEET in wb.sheetnames:
        del wb[MTOATD_SHEET]

    # --- indeks D & W per currency per tanggal
    idx = spec.Indeks()
    tanpa_tanggal = 0
    for rec in report["records"]:
        cur = norm(rec.get("Currency"))
        if not cur:
            continue
        settle = as_date(rec.get("Settlement Date"))
        usd = to_float(rec.get("USD")) or 0.0
        tx = to_float(rec.get("Transaction")) or 0.0
        if rec["Type"] == "withdrawal":
            selesai = as_date(rec.get("Completed Date")) or as_date(rec.get("Date"))
            if selesai is None and settle is None:
                tanpa_tanggal += 1
                continue
            idx.tambah_withdrawal(cur, selesai, settle, usd, tx,
                                  to_float(rec.get("Charges")) or 0.0,
                                  rec.get("Status"), rec.get("Source Name"))
        else:
            paid = as_date(rec.get("Paid Date")) or as_date(rec.get("Date"))
            if paid is None and settle is None:
                tanpa_tanggal += 1
                continue
            idx.tambah_deposit(cur, paid, settle, usd, tx)

    # Baris `refuse` untuk aturan H+1 -- daftar_tanggal=False supaya tidak
    # menciptakan blok tanggal baru; dia cuma menumpang di blok yang sudah ada.
    n_ref = 0
    for rec in report.get("refuse_mtoatd") or []:
        if not rec["cur"] or (rec["completed"] is None and rec["settle"] is None):
            continue
        idx.tambah_withdrawal(rec["cur"], rec["completed"], rec["settle"],
                              rec["usd"] or 0.0, rec["tx"] or 0.0, rec["charges"] or 0.0,
                              "refuse", rec["source"], daftar_tanggal=False)
        n_ref += 1
    if n_ref:
        print(f"         MTOATD: {n_ref:,} baris 'refuse' ikut disertakan untuk aturan "
              f"H+1 (Settlement = t, Completed = t+1) di baris MT4出金 / Wallet出金 / "
              f"CRM出金（原币种）")

    if not idx.tanggal:
        return 0, 0
    if tanpa_tanggal:
        print(f"         MTOATD: {tanpa_tanggal} baris tanpa tanggal sama sekali -> dilewati")

    tanggal = sorted(idx.tanggal)
    # kolom currency: urutan sheet yang sudah beredar dulu, sisanya di belakang
    extra = sorted(c for c in idx.currencies if c not in spec.URUT_CURRENCY)
    urut_cur = list(spec.URUT_CURRENCY) + extra
    kol_cur = {c: spec.KOL_CUR_AWAL + i for i, c in enumerate(urut_cur)}
    kol_total = spec.KOL_CUR_AWAL + len(urut_cur)
    # blok akumulasi di kanan: 1 kolom kosong, lalu label + tanggal + currency
    kol_ak_label = kol_total + 2
    kol_ak_cur = {c: kol_ak_label + 2 + i for i, c in enumerate(urut_cur)}
    kol_ak_total = kol_ak_label + 2 + len(urut_cur)
    pakai_ak = MTOATD_BLOK_AKUMULASI

    ws = wb.create_sheet(MTOATD_SHEET)
    ws.sheet_properties.tabColor = "375623"

    def label_bulan(d):
        """Label baris 月累计 mengikuti BULAN blok itu, bukan bulan pertama data."""
        return (f"{d.year}.{d.month}\u6708\u7d2f\u8ba1 "
                f"MT4+\u9322\u5305\u6de8\u5165\u91d1 (USD)")

    ws.cell(1, 1, "MTOATD - daily reconciliation").font = Font(bold=True, size=14)
    ws.cell(1, 4, "Every row that has a formula in the source workbook is computed here "
                  "from sheet D and W. Rows that have no formula there (manual breakdowns "
                  "such as voided orders, rounding differences, Wallet deposit) are left "
                  "EMPTY on purpose - grey. Green = computed. "
                  "Right-hand block = month-to-date accumulation.").font = \
        Font(italic=True, size=9, color="808080")

    hijau = PatternFill("solid", fgColor="E2EFDA")
    abu = PatternFill("solid", fgColor="F2F2F2")
    biru = PatternFill("solid", fgColor="DDEBF7")
    hijau_tua = PatternFill("solid", fgColor="A9D08E")

    akum_net = defaultdict(float)      # cur -> akumulasi 本日 (untuk baris 月累计)
    akum_ak = defaultdict(float)       # (kunci, cur) -> akumulasi blok 当月累计
    n_sel = 0
    r0 = 3
    bulan_akum = None                  # bulan yang sedang diakumulasi
    for d in tanggal:
        # 月累计 dan blok 当月累计 artinya "sejak awal BULAN INI". Begitu bulannya
        # ganti, akumulasinya HARUS mulai dari nol lagi -- kalau tidak, blok bulan
        # berikutnya ikut membawa angka bulan sebelumnya. Satu transaksi nyasar
        # ke bulan lain (mis. Completed Date 1 Juli di laporan Juni) sudah cukup
        # untuk memunculkan blok yang angkanya salah.
        if bulan_akum != (d.year, d.month):
            bulan_akum = (d.year, d.month)
            akum_net.clear()
            akum_ak.clear()
        label_cum = label_bulan(d)
        nilai = {c: spec.hitung(idx, c, d) for c in urut_cur}
        for c in urut_cur:
            akum_net[c] += nilai[c]["net_hari"]
            nilai[c]["net_bulan"] = akum_net[c]

        for off, label, kat, jenis, kunci in spec.BARIS:
            r = r0 + off
            if jenis == "blank":
                continue
            if kat:
                ws.cell(r, 1, kat).font = Font(bold=True)
            if jenis == "header":
                for kl, kd, kc, kt in ((1, 2, kol_cur, kol_total),
                                       (kol_ak_label, kol_ak_label + 1, kol_ak_cur, kol_ak_total)):
                    if kl != 1 and not pakai_ak:
                        break
                    ws.cell(r, kl, "\u65e5\u671f\uff1a").font = Font(bold=True)
                    sel = ws.cell(r, kd, d)
                    sel.number_format = "yyyy-mm-dd"
                    sel.font = Font(bold=True)
                    for cur in urut_cur:
                        x = ws.cell(r, kc[cur], cur)
                        x.font = Font(bold=True)
                        x.fill = biru
                        x.alignment = Alignment(horizontal="center")
                    x = ws.cell(r, kt, "\u5408\u8a08")
                    x.font = Font(bold=True)
                    x.fill = biru
                if pakai_ak:
                    ws.cell(r - 1, kol_ak_cur[urut_cur[0]], "\u5f53\u6708\u7d2f\u8ba1").font = \
                        Font(bold=True, size=11)
                continue

            teks = label_cum if jenis == "cum" else label
            sel = ws.cell(r, 2, teks)
            sel.font = Font(bold=True)
            if label in spec.EN:
                sel.comment = Comment(spec.EN[label], "mtoatd", height=60, width=300)
            if pakai_ak:
                ws.cell(r, kol_ak_label + 1, teks).font = Font(bold=True)

            for cur in urut_cur:
                v = nilai[cur].get(kunci) if kunci else None
                cell = ws.cell(r, kol_cur[cur])
                cell.number_format = FMT_ACC
                if jenis == "manual" or v is None:
                    cell.fill = abu                     # tidak ada rumusnya -> kosong
                else:
                    cell.value = v
                    cell.fill = hijau_tua if jenis in ("calc", "cum") and \
                        kunci in ("net_hari", "net_bulan") else hijau
                    n_sel += 1
                if pakai_ak:
                    ca = ws.cell(r, kol_ak_cur[cur])
                    ca.number_format = FMT_ACC
                    if jenis == "manual" or v is None:
                        ca.fill = abu
                    else:
                        if kunci == "net_bulan":
                            ca.value = v                # sudah akumulasi
                        else:
                            akum_ak[(kunci, cur)] += v
                            ca.value = akum_ak[(kunci, cur)]
                        ca.fill = hijau

            # kolom 合計 = jumlah antar currency
            for kt, kc in ((kol_total, kol_cur),) + (((kol_ak_total, kol_ak_cur),) if pakai_ak else ()):
                angka = [ws.cell(r, kc[cur]).value for cur in urut_cur]
                angka = [x for x in angka if isinstance(x, (int, float))]
                sel = ws.cell(r, kt)
                sel.number_format = FMT_ACC
                sel.font = Font(bold=True)
                if angka:
                    sel.value = sum(angka)
        r0 += spec.TINGGI_BLOK

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 32
    for cur in urut_cur:
        ws.column_dimensions[openpyxl.utils.get_column_letter(kol_cur[cur])].width = 16
    ws.column_dimensions[openpyxl.utils.get_column_letter(kol_total)].width = 18
    if pakai_ak:
        ws.column_dimensions[openpyxl.utils.get_column_letter(kol_ak_label + 1)].width = 32
        for cur in urut_cur:
            ws.column_dimensions[openpyxl.utils.get_column_letter(kol_ak_cur[cur])].width = 16
        ws.column_dimensions[openpyxl.utils.get_column_letter(kol_ak_total)].width = 18
    ws.freeze_panes = "C1"
    report["mtoatd_sel"] = n_sel
    return len(tanggal), len(urut_cur)


def write_dw_report_sheet(wb, report, periode):
    """Sheet 'D&W Report': ringkasan Currency x Payment Gateway, layout sama
    dengan report D&W yang sudah beredar (Deposit di atas, Withdrawal di bawah)."""
    from openpyxl.styles import Alignment

    if REPORT_SHEET in wb.sheetnames:
        del wb[REPORT_SHEET]
    ws = wb.create_sheet(REPORT_SHEET, 0)
    ws.sheet_properties.tabColor = "2E7D32"

    KOL_D = ["Currency", "Payment Gateway", "CRM USD ", "TRANSACTIONS", "Xero USD ",
             "Total FE Gain/-Loss (USD)", "CRM Average Rate", "Xero average Rate",
             "Recorded in Xero", "Rounding Adjustment"]
    KOL_W = ["Currency", "Payment Gateway", "CRM USD ", "Handling Charges Income",
             "Transactions", "Xero USD ", "Total FE Gain/-Loss (USD)", "CRM average",
             "Xero average", "Recorded in Xero", "Rounding Adjustment"]

    for kol, w in zip("ABCDEFGHIJK",
                      (10, 16.7, 13.3, 18, 17, 14.3, 16.8, 16.7, 15.2, 19, 19)):
        ws.column_dimensions[kol].width = w

    ws["A1"] = REPORT_TITLE
    ws["A1"].font = Font(bold=True, size=16)

    r = 3
    for jenis in ("deposit", "withdrawal"):
        data = [((c, g), v) for (k, c, g), v in report["pivot"].items() if k == jenis]
        r += 1
        ws.cell(r, 1, REPORT_SECTION_TITLE[jenis]).font = Font(bold=True, size=16)
        r += 1
        sel = ws.cell(r, 1, periode)
        sel.font = Font(bold=True, size=16)
        sel.number_format = "mmm-yy"
        if jenis == "deposit":
            r += 2
            ws.cell(r, 1, report["date_col_label"].get("deposit", "Date")).font = Font(size=12)
            ws.cell(r, 2, "(Multiple Items)").font = Font(size=12)
            r += 1
        else:
            r += 1
        ws.cell(r, 1, "Payment Gateway").font = Font(bold=True, size=12)

        kolom = KOL_D if jenis == "deposit" else KOL_W
        n_data = 5 if jenis == "deposit" else 6      # kolom terakhir yg berwarna biru-kelabu
        r += 1
        baris_hdr = r
        for c, t in enumerate(kolom, start=1):
            sel = ws.cell(r, c, t)
            sel.font = Font(bold=True)
            sel.fill = PatternFill("solid",
                                   fgColor=RPT_FILL_HEAD if c <= n_data else RPT_FILL_HEAD2)
            sel.alignment = Alignment(wrap_text=True, vertical="bottom")

        cur_terakhir = None
        tot = {"crm": 0.0, "tx": 0.0, "xu": 0.0, "ch": 0.0}
        for cur, gw, v in _rpt_urut(data):
            r += 1
            crm, tx, xu, ch = v["crm_usd"], v["tx"], v["xero_usd"], v["charges"]
            tot["crm"] += crm; tot["tx"] += tx; tot["xu"] += xu; tot["ch"] += ch

            if cur != cur_terakhir:
                sel = ws.cell(r, 1, cur)
                sel.font = Font(bold=True)
                sel.fill = PatternFill("solid", fgColor=RPT_FILL_CUR)
                cur_terakhir = cur
            ws.cell(r, 2, gw)

            if jenis == "deposit":
                nilai = [crm, tx, xu, xu - crm,
                         (tx / crm if crm else None), (tx / xu if xu else None),
                         None, -xu]
            else:
                nilai = [crm, ch, tx, xu, crm - xu,
                         (tx / crm if crm else None), (tx / xu if xu else None),
                         None, xu]
            for i, val in enumerate(nilai):
                c = 3 + i
                sel = ws.cell(r, c, val)
                sel.number_format = FMT_ACC5 if kolom[c - 1].lower().endswith(
                    ("rate", "average")) else FMT_ACC
                if c > n_data:
                    sel.fill = PatternFill("solid", fgColor=RPT_FILL_CALC)

        # ---- Grand Total
        r += 1
        for c in range(1, len(kolom) + 1):
            sel = ws.cell(r, c)
            sel.font = Font(bold=True)
            sel.fill = PatternFill("solid",
                                   fgColor=RPT_FILL_HEAD if c <= n_data else RPT_FILL_HEAD2)
        ws.cell(r, 1, "Grand Total")
        crm_avg = (tot["tx"] / tot["crm"]) if tot["crm"] else None
        xero_avg = (tot["tx"] / tot["xu"]) if tot["xu"] else None
        # 'Recorded in Xero' diisi manual -> jumlahkan kalau memang sudah ada isinya
        c_rec = kolom.index("Recorded in Xero") + 1
        isi_rec = [ws.cell(rr, c_rec).value for rr in range(baris_hdr + 1, r)]
        rec_tot = sum(x for x in isi_rec if isinstance(x, (int, float))) \
            if any(isinstance(x, (int, float)) for x in isi_rec) else None
        if jenis == "deposit":
            #  C          D         E         F                    G        H         I     J
            gt = [tot["crm"], tot["tx"], tot["xu"], tot["xu"] - tot["crm"],
                  crm_avg, xero_avg, rec_tot, -tot["xu"]]
            kol_rate = (7, 8)          # G, H
        else:
            #  C          D        E         F         G                    H        I         J     K
            gt = [tot["crm"], tot["ch"], tot["tx"], tot["xu"], tot["crm"] - tot["xu"],
                  crm_avg, xero_avg, rec_tot, tot["xu"]]
            kol_rate = (8, 9)          # H, I
        for i, val in enumerate(gt):
            c = 3 + i
            sel = ws.cell(r, c, val)
            sel.number_format = FMT_ACC5 if c in kol_rate else FMT_ACC
            sel.font = Font(bold=True)
        ws.auto_filter.ref = None
        r += 2
    return ws


def write_detail_sheet(wb, report):
    """Sheet 'D&W Detail': gabungan semua baris sheet D + W + kolom hitungan."""
    from openpyxl.styles import Alignment

    if DETAIL_SHEET in wb.sheetnames:
        del wb[DETAIL_SHEET]
    ws = wb.create_sheet(DETAIL_SHEET)
    ws.sheet_properties.tabColor = "1F4E79"

    kolom = [n for n, _ in REPORT_COLUMNS] + NEW_COLUMNS
    if REPORT_AUDIT_COLUMNS:
        kolom += ["Type", "Source Sheet", "Date Source", "Data Status"]

    for c, nama in enumerate(kolom, start=1):
        cell = ws.cell(1, c, nama)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E7D32")
        if nama in LEGENDA:
            cell.comment = Comment(LEGENDA[nama], "hitung_dw.py", height=150, width=320)
    ws.freeze_panes = "A2"

    fmt = {"Handling Fee": FMT_FEE, "Xero Rate": FMT_RATE,
           "Xero USD": FMT_MONEY, "Forex Gain/Loss": FMT_MONEY}
    tgl_kol = {"Date", "Settlement Date"}

    # Sheet 'D' tidak punya kolom Status sama sekali (deposit tidak lewat
    # approve/refuse), jadi kolom Status baris deposit KOSONG. Di filter Excel itu
    # muncul sebagai "(Blanks)" dan terlihat seperti data gagal yang lolos, padahal
    # semua baris withdrawal di sini sudah pasti 'finish'. Diberi label eksplisit
    # supaya tidak ada sel kosong di kolom itu.
    STATUS_KOSONG = "n/a (deposit)"

    r = 1
    for rec in report["records"]:
        r += 1
        for c, nama in enumerate(kolom, start=1):
            nilai = rec.get(nama)
            if nama == "Status" and nilai in (None, ""):
                nilai = STATUS_KOSONG
            cell = ws.cell(r, c, nilai)
            if nama in fmt:
                cell.number_format = fmt[nama]
            elif nama in tgl_kol and rec.get(nama) is not None:
                cell.number_format = "yyyy-mm-dd"
            if nama == "Handling Fee" and rec["_fee_missing"]:
                cell.fill, cell.font = FILL_MISSING, FONT_MISSING
            elif nama in ("Xero Rate", "Xero USD", "Forex Gain/Loss") and rec["_fx_missing"]:
                cell.fill, cell.font = FILL_MISSING, FONT_MISSING
            elif nama == "Xero Rate" and rec["_fx_fallback"]:
                cell.fill = FILL_WARN
            elif nama == "Data Status" and rec["Data Status"] != "OK":
                cell.fill = FILL_WARN if not (rec["_fee_missing"] or rec["_fx_missing"]) else FILL_MISSING
                cell.alignment = Alignment(wrap_text=False)

    lebar = {"Data Status": 52, "Sales Name": 22, "Sales Dept": 22, "Sales Area": 24,
             "Payment Gateway": 20, "Date Source": 15, "Source Sheet": 13, "Hash": 26,
             "Reference": 20, "Settlement Date": 15, "Date": 12}
    for c, nama in enumerate(kolom, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = \
            lebar.get(nama, max(11, len(nama) + 3))
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(kolom))}{max(r, 2)}"
    return r - 1


def write_missing_sheet(wb, report):
    """Sheet 'Missing Data': rincian tanggal kurs & rate fee yang belum ada."""
    from openpyxl.styles import Alignment, Border, Side

    if MISSING_SHEET in wb.sheetnames:
        del wb[MISSING_SHEET]
    ws = wb.create_sheet(MISSING_SHEET)
    ws.sheet_properties.tabColor = "C00000"
    thin = Side(style="thin", color="D0D0D0")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(vertical="top", wrap_text=True)

    for col, w in zip("ABCDEFG", (3, 16, 22, 15, 16, 16, 46)):
        ws.column_dimensions[col].width = w

    r = 2
    ws.cell(r, 2, "Missing Data Report").font = Font(bold=True, size=16)
    r += 1
    ws.cell(r, 2, "Every row below is data that is not in the workbook yet. Until it is "
                  "added, the affected cells stay RED (nothing calculated) or "
                  "YELLOW (approximated).").font = Font(italic=True, size=9, color="808080")

    # Peringatan paling penting: TIDAK ADA satu pun baris transaksi terbaca.
    # Tanpa ini file hasilnya kelihatan normal padahal kosong -- sheet MTOATD /
    # Channel Balance / J Wallet malah tidak dibuat sama sekali.
    if not report["records"]:
        r += 2
        sel = ws.cell(r, 2, "NO TRANSACTION ROWS WERE FOUND - THIS RESULT IS EMPTY")
        sel.font = Font(bold=True, size=13, color="9C0006")
        sel.fill = FILL_MISSING
        r += 1
        ws.cell(r, 2, "Sheet 'D' and 'W' of the file you uploaded have no data rows, so "
                      "nothing could be calculated and the sheets 'MTOATD', "
                      "'Channel Balance' and 'J Wallet (calc)' were not created at all. "
                      "Paste the deposit rows into sheet 'D' starting at row 2 and the "
                      "withdrawal rows into sheet 'W' starting at row 2, keep the header "
                      "row as it is, then upload again.").font = \
            Font(italic=True, size=10, color="9C0006")

    # ---- 0. bulan laporan + baris yang dibuang karenanya
    r += 2
    ws.cell(r, 2, "0. REPORT MONTH").font = Font(bold=True, size=12, color="4472C4")
    r += 1
    if not PERIODE_FILTER:
        ws.cell(r, 2, "No report month was chosen, so EVERY date found in the file was "
                      "calculated. A back-office export normally carries a tail of the "
                      "month before and after (the export filters on Apply Date while "
                      "this report uses Paid Date / Completed Date), so the totals below "
                      "may cover more than one month.").font = \
            Font(italic=True, size=10, color="C00000")
        ws.row_dimensions[r].height = 44
        r += 1
    else:
        ws.cell(r, 2, f"Report month: {nama_periode()}. Only rows whose date falls in "
                      f"this month are counted - in D, in W and in the Fund Transfer "
                      f"Table. The date used is the same one the whole report uses: "
                      f"Paid Date for deposits, Completed Date for withdrawals.").font = \
            Font(size=10, color="008000")
        ws.row_dimensions[r].height = 30
        r += 1
        buang = report.get("luar_periode") or {}
        if not buang:
            ws.cell(r, 2, "None - every row in the file already fell inside this "
                          "month.").font = Font(italic=True, color="008000")
            r += 1
        else:
            _header(ws, r, ["Sheet", "Month of the row", "Rows dropped", "USD dropped",
                            "", "Why"], "4472C4")
            for (sh, bln), jml in sorted(buang.items()):
                r += 1
                usd = (report.get("luar_periode_usd") or {}).get((sh, bln))
                isi = [sh, bln, jml, usd, "",
                       f"Outside {nama_periode()} - not part of this month's report"]
                for c, v in enumerate(isi, start=2):
                    cell = ws.cell(r, c, v)
                    cell.border = box
                    cell.alignment = wrap
                    if c == 5 and v:
                        cell.number_format = FMT_MONEY
                    if c == 4:
                        cell.fill = FILL_SKIP
            r += 1
            ws.cell(r, 2, f"TOTAL: {sum(buang.values()):,} rows dropped because they are "
                          f"not in {nama_periode()}. This is expected, not an error - "
                          f"their calculated columns in D / W are left blank and shaded "
                          f"grey.").font = Font(bold=True)
            ws.row_dimensions[r].height = 30

    # ---- 1. tanggal kurs yang tidak ada di XERO
    r += 2
    ws.cell(r, 2, "1. MISSING EXCHANGE RATE DATES (sheet 'XERO')").font = \
        Font(bold=True, size=12, color="C00000")
    r += 1
    if not report["fx_detail"]:
        ws.cell(r, 2, "None - every transaction date was found in the XERO sheet.").font = \
            Font(italic=True, color="008000")
        r += 1
    else:
        _header(ws, r, ["Currency", "Date not in XERO", "Rows affected",
                        "Rate used instead", "Days earlier", "What to do"], "C00000")
        for det in sorted(report["fx_detail"].values(),
                          key=lambda d: (-d["rows"], d["currency"], str(d["missing_date"]))):
            r += 1
            hari = ((det["missing_date"] - det["used_date"]).days
                    if det["missing_date"] and det["used_date"] else None)
            isi = [det["currency"], det["missing_date"], det["rows"],
                   det["used_date"], hari,
                   f"Add {det['currency']} rate for {det['missing_date']} to the XERO sheet"]
            for c, v in enumerate(isi, start=2):
                cell = ws.cell(r, c, v)
                cell.border = box
                cell.alignment = wrap
                if c in (3, 5) and v is not None:
                    cell.number_format = "yyyy-mm-dd"
                if c == 2:
                    cell.fill = FILL_WARN
        r += 1
        tot = sum(d["rows"] for d in report["fx_detail"].values())
        ws.cell(r, 2, f"TOTAL: {len(report['fx_detail'])} currency/date combinations, "
                      f"{tot:,} rows approximated").font = Font(bold=True)

    # ---- 2. rate fee yang belum ada
    r += 2
    ws.cell(r, 2, "2. MISSING FEE RATES (sheet 'D&W FEE')").font = \
        Font(bold=True, size=12, color="C00000")
    r += 1
    if not report["fee_missing"]:
        ws.cell(r, 2, "None - every Currency + Payment Gateway had a fee rate.").font = \
            Font(italic=True, color="008000")
        r += 1
    else:
        _header(ws, r, ["Currency", "Payment Gateway", "Rows affected", "", "",
                        "What to do"], "C00000")
        for (cur, gw), n in sorted(report["fee_missing"].items(), key=lambda x: -x[1]):
            r += 1
            isi = [cur, gw, n, "", "",
                   f"Add row '{cur} | {gw}' with its Deposit and Withdrawal rate "
                   f"to the 'D&W FEE' sheet"]
            for c, v in enumerate(isi, start=2):
                cell = ws.cell(r, c, v)
                cell.border = box
                cell.alignment = wrap
                if c == 2:
                    cell.fill = FILL_MISSING
                    cell.font = FONT_MISSING
        r += 1
        ws.cell(r, 2, f"TOTAL: {len(report['fee_missing'])} combinations, "
                      f"{sum(report['fee_missing'].values()):,} rows with no Handling Fee"
                      ).font = Font(bold=True)

    # ---- 2b. rate yang perlu dikonfirmasi
    r += 2
    ws.cell(r, 2, "2b. FEE RATES THAT NEED CONFIRMATION").font = \
        Font(bold=True, size=12, color="C00000")
    r += 1
    dipakai = {(c, g) for c, g, *_ in
               ((k[0], k[1]) + tuple() for k in report["fee_match"])} if False else set()
    for kunci in report["fee_match"]:
        dipakai.add((kunci[0], kunci[1]))
    perlu = [(c, g, txt) for (c, g), txt in PERLU_KONFIRMASI.items()]
    if not perlu:
        ws.cell(r, 2, "None.").font = Font(italic=True, color="008000")
    else:
        _header(ws, r, ["Currency", "Payment Gateway", "Rows using it", "", "",
                        "Why it needs confirming"], "C00000")
        for c_, g_, txt in perlu:
            r += 1
            n = sum(v for k, v in report["fee_match"].items()
                    if k[0] == norm(c_) and kunci_gw(k[1]) == kunci_gw(g_))
            for c, v in enumerate([c_, g_, n or "none in this file", "", "", txt], start=2):
                cell = ws.cell(r, c, v)
                cell.border = box
                cell.alignment = wrap
                if c == 2:
                    cell.fill = FILL_WARN
            ws.row_dimensions[r].height = 84

    # ---- 3. masalah data lain
    r += 2
    ws.cell(r, 2, "3. INPUT SHEETS THAT ARE MISSING OR EMPTY").font = \
        Font(bold=True, size=12, color="C00000")
    r += 1
    kurang = report.get("input_kurang") or []
    if not kurang:
        ws.cell(r, 2, "None - every input sheet had data.").font = \
            Font(italic=True, color="008000")
    else:
        _header(ws, r, ["Sheet", "", "What is wrong", "", "", "What it affects / what to do"],
                "C00000")
        for nama, sebab, akibat, aksi in kurang:
            r += 1
            for c, v in enumerate([nama, "", sebab, "", "", akibat + " -> " + aksi], start=2):
                cell = ws.cell(r, c, v)
                cell.border = box
                cell.alignment = wrap
            ws.row_dimensions[r].height = 58
    r += 2

    ws.cell(r, 2, "4. OTHER DATA ISSUES").font = Font(bold=True, size=12, color="C00000")
    r += 1
    lain = Counter()
    for rec in report["records"]:
        for bagian in rec["Data Status"].split("; "):
            if bagian.startswith("DATA:"):
                lain[bagian] += 1
    if not lain:
        ws.cell(r, 2, "None - no empty gateway, amount or date found.").font = \
            Font(italic=True, color="008000")
    else:
        _header(ws, r, ["Issue", "", "Rows affected", "", "", "What to do"], "C00000")
        for isu, n in lain.most_common():
            r += 1
            for c, v in enumerate([isu.replace("DATA: ", ""), "", n, "", "",
                                   "Fix in the source export (back office), "
                                   "not in the fee table"], start=2):
                cell = ws.cell(r, c, v)
                cell.border = box
                cell.alignment = wrap
    return ws


def write_legend_sheet(wb, date_cols_used):
    """Bikin/timpa sheet 'Legend' berisi penjelasan kolom & warna (English)."""
    from openpyxl.styles import Alignment, Border, Side

    if LEGEND_SHEET in wb.sheetnames:
        del wb[LEGEND_SHEET]
    ws = wb.create_sheet(LEGEND_SHEET)
    ws.sheet_properties.tabColor = "4472C4"

    thin = Side(style="thin", color="D0D0D0")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(vertical="top", wrap_text=True)

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 62
    ws.column_dimensions["D"].width = 52

    r = 2
    ws.cell(r, 2, "Calculated Columns & Colour Legend").font = Font(bold=True, size=16)
    r += 1
    ws.cell(r, 2, "Generated automatically by hitung_dw.py. Do not edit this sheet - "
                  "it is rewritten every time the calculation runs.").font = Font(italic=True, size=9, color="808080")
    # Cap waktu: user sempat memeriksa file hasil LAMA dan mengira aturan baru tidak
    # jalan. Dengan stempel ini, versi file langsung ketahuan.
    r += 1
    aturan = ("only rows with Status = 'finish' are counted"
              if HANYA_STATUS_FINISH else "ALL rows are counted regardless of Status")
    bulan = (f"REPORT MONTH {nama_periode()} - rows dated outside it are dropped"
             if PERIODE_FILTER else
             "NO report month was chosen - every date in the file is counted")
    ws.cell(r, 2, f"GENERATED {datetime.datetime.now():%Y-%m-%d %H:%M} - {bulan}; "
                  f"{aturan}. "
                  f"If a result file does not carry this line, it is an OLD file.").font = \
        Font(bold=True, size=10, color="C00000")
    ws.row_dimensions[r].height = 28

    # --- bagian 1: kolom
    r += 2
    ws.cell(r, 2, "CALCULATED COLUMNS").font = Font(bold=True, size=12, color="4472C4")
    r += 1
    for c, t in ((2, "Column"), (3, "How it is calculated"), (4, "Notes")):
        cell = ws.cell(r, c, t)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.border = box
    rows = [
        ("Handling Fee",
         "Transaction x fee rate from the 'D&W FEE' sheet, matched on "
         "Currency + Payment Gateway.",
         "Deposit rows use the Deposit rate; withdrawal rows use the Withdrawal rate. "
         "Deposit vs withdrawal is read from the 'Source Name' column of each row."),
        ("Xero Rate",
         "Exchange rate from the 'XERO' sheet, looked up by transaction date + currency.",
         "Quoted as local currency per 1 USD (e.g. VND 26,317.2 = 1 USD). "
         + " ".join(f"Sheet '{k}' uses the '{v}' column." for k, v in date_cols_used.items())),
        ("Xero USD",
         "Transaction / Xero Rate",
         "The transaction amount restated in USD at the XERO rate."),
        ("Forex Gain/Loss",
         "Deposit rows:     Xero USD - USD\n"
         "Withdrawal rows:  USD - Xero USD",
         "The sign is flipped for withdrawals because the money is going out, "
         "so a deposit gain and a withdrawal gain point the same way."),
    ]
    for name, formula, note in rows:
        r += 1
        for c, t in ((2, name), (3, formula), (4, note)):
            cell = ws.cell(r, c, t)
            cell.border = box
            cell.alignment = wrap
            if c == 2:
                cell.font = Font(bold=True)
        ws.row_dimensions[r].height = 34

    # --- bagian 2: warna
    r += 2
    ws.cell(r, 2, "CELL COLOUR LEGEND").font = Font(bold=True, size=12, color="4472C4")
    r += 1
    for c, t in ((2, "Colour"), (3, "What it means"), (4, "How to fix it")):
        cell = ws.cell(r, c, t)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.border = box
    colours = [
        (FILL_MISSING, FONT_MISSING, "RED - empty",
         "No fee rate exists yet for this Currency + Payment Gateway combination, "
         "so 'Handling Fee' was deliberately left EMPTY instead of 0. "
         "An empty cell is excluded from SUM, so totals are never silently understated.",
         "Add the missing Currency + Payment Gateway row to the 'D&W FEE' sheet "
         "with its Deposit and Withdrawal rate, then run the calculation again."),
        (FILL_MISSING, FONT_MISSING, "RED - in a rate column",
         "In 'Xero Rate', 'Xero USD' or 'Forex Gain/Loss': no exchange rate could be "
         "found at all for that date and currency, so nothing could be calculated.",
         "Add the currency column and/or the missing dates to the 'XERO' sheet, "
         "then run the calculation again."),
        (FILL_WARN, Font(), "YELLOW",
         "In 'Xero Rate': the transaction date is not present in the 'XERO' sheet, "
         "so the rate of the nearest earlier date was used instead. The figure is "
         "usable but it is an approximation, not the rate of that exact day.",
         "Extend the 'XERO' sheet so it covers every transaction date, "
         "then run the calculation again."),
        (PatternFill(fill_type=None), Font(), "NO FILL",
         "Normal. The value was calculated from an exact match on both the fee rate "
         "and the exchange rate of that exact date.",
         "Nothing to do."),
    ]
    for fill, font, label, meaning, fix in colours:
        r += 1
        swatch = ws.cell(r, 2, label)
        swatch.fill = fill
        swatch.font = font if font.bold else Font(bold=True)
        swatch.border = box
        swatch.alignment = wrap
        for c, t in ((3, meaning), (4, fix)):
            cell = ws.cell(r, c, t)
            cell.border = box
            cell.alignment = wrap
        ws.row_dimensions[r].height = 58

    # --- bagian: bulan laporan
    r += 2
    ws.cell(r, 2, "REPORT MONTH").font = Font(bold=True, size=12, color="C00000")
    r += 1
    if PERIODE_FILTER:
        ws.cell(r, 2, f"This file was generated for {nama_periode()}. Any row dated "
                      f"outside that month was dropped completely - in sheet D, in "
                      f"sheet W and in the Fund Transfer Table - so no total, balance "
                      f"or MTOATD block on any sheet contains money from another "
                      f"month. A back-office export normally carries a tail either "
                      f"side (an export filtered on Apply Date can still hold rows "
                      f"whose Paid Date or Completed Date lands in the next month); "
                      f"those tails are what gets removed. The date each row is judged "
                      f"on is the same one the whole report uses: Paid Date for "
                      f"deposits, Completed Date for withdrawals. Rows with no readable "
                      f"date at all are dropped too, and counted separately as "
                      f"'(no date)'. The exact numbers are on sheet "
                      f"'{MISSING_SHEET}', section 0.").font = \
            Font(size=10, color="C00000")
        ws.row_dimensions[r].height = 88
        r += 1
        ws.cell(r, 2, "In sheets D and W the dropped rows are still there, but their "
                      "four calculated columns are blank and shaded grey. Filter those "
                      "columns for blanks to see exactly which rows were left out.").font = \
            Font(italic=True, size=10, color="808080")
        ws.row_dimensions[r].height = 30
    else:
        ws.cell(r, 2, "NO report month was chosen for this file, so every date found "
                      "in it was calculated. If the upload covered more than one month "
                      "(a June export commonly runs from late May to early July), the "
                      "totals, the balances and the MTOATD date blocks span all of it. "
                      "Choose the report month on the upload page - or pass "
                      "--period YYYY-MM - to get a single clean month.").font = \
            Font(size=10, color="C00000")
        ws.row_dimensions[r].height = 44

    # --- bagian: aturan status
    r += 2
    ws.cell(r, 2, "ONLY 'FINISH' ROWS ARE COUNTED").font = \
        Font(bold=True, size=12, color="C00000")
    r += 1
    ws.cell(r, 2, "Every row in sheet D and W whose Status is not 'finish' is dropped "
                  "completely - it is not counted in any sheet, total or variance. "
                  "That covers rows marked 'refuse' AND rows where the Status cell was "
                  "left empty (those carry remarks such as 'Application Failed, please "
                  "withdraw...'). A withdrawal that was refused never left the account, "
                  "so counting it would overstate money out. The number of rows dropped "
                  "is always printed when the calculation runs.").font = \
        Font(size=10, color="C00000")
    ws.row_dimensions[r].height = 44
    r += 1
    ws.cell(r, 2, "CONFIRMED by the DPM Malaysia team, 3 September 2026, on the treatment of "
                  "a blank Status cell: \"Blank status is same as Refuse, hence do not "
                  "include in any calculations.\" That is exactly what this report does. "
                  "Note that a SUMIFS criterion of \"<>refuse\" does NOT do this - in Excel "
                  "that also matches empty cells, so a formula written that way still counts "
                  "the blank-Status rows as money paid out.").font = \
        Font(size=10, color="008000")
    ws.row_dimensions[r].height = 44
    r += 1
    ws.cell(r, 2, "Note on sheet 'D&W Detail': the deposit rows show an EMPTY Status "
                  "column. That is normal - sheet 'D' has no Status column at all, "
                  "because a deposit has no approve/refuse step. Only withdrawals carry "
                  "a status, and every withdrawal row you see here is 'finish'. Filter "
                  "the 'Type' column to tell deposits and withdrawals apart.").font = \
        Font(size=10, color="C00000")
    ws.row_dimensions[r].height = 40
    r += 1
    ws.cell(r, 2, "HEADS UP for anyone comparing against the older Excel report: that "
                  "report filters 'refuse' only on the TD rows. Its MT4 and CRM "
                  "withdrawal rows include refused withdrawals, so its "
                  "'CRM - TD Withdrawal Variance' contains those refusals as if they "
                  "were a timing difference. Figures here will therefore differ on "
                  "MT4 Withdrawal, CRM Withdrawal and the variances derived from them.").font = \
        Font(italic=True, size=10, color="808080")
    ws.row_dimensions[r].height = 44

    # --- bagian 3: sheet MTOATD
    r += 2
    ws.cell(r, 2, "SHEET 'MTOATD'").font = Font(bold=True, size=12, color="4472C4")
    r += 1
    for c, t in ((2, "Item"), (3, "What it means"), (4, "Notes")):
        cell = ws.cell(r, c, t)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.border = box
    mtoatd_rows = [
        ("GREEN cell",
         "Computed from sheet D and W. Every row that has a formula in the original "
         "MTOATD sheet is reproduced here, so nothing has to be typed in by hand.",
         "Hover over a row label in column B to see the exact rule for that row."),
        ("GREY - empty",
         "That row has NO formula in the original MTOATD sheet - it is a manual "
         "breakdown (voided orders, rounding differences, Wallet deposit, and so on). "
         "It is left EMPTY on purpose rather than filled with 0.",
         "Type into it if you want to record a breakdown. Nothing recalculates from it."),
        ("Date basis",
         "MT4 / CRM / \u5f53\u6708 rows use SETTLEMENT DATE.\n"
         "TD deposit and \u5b9e\u6536 use PAID DATE.\n"
         "TD withdrawal and \u5b9e\u51fa use COMPLETED DATE.",
         "A day's block only appears for dates that exist in the transaction data."),
        ("Right-hand block",
         "\u5f53\u6708\u7d2f\u8ba1 - the same rows accumulated from the first day of the month "
         "up to and including that day.",
         "In the original workbook a few of these cells were dragged inconsistently; "
         "here every row accumulates the same way."),
        ("Known quirks kept",
         "Three rows in the original sheet were edited only in the USDT / USD columns: "
         "\u5b9e\u51fa\uff08\u539f\u5e01\u79cd\uff09 filters out 'refuse' for USDT only, "
         "\u2014\u2014\u624b\u7eed\u8d39/\u8f6c\u8d26\u8d39 exists only for USDT and USD, and the last "
         "\u2014\u2014\u8c03\u4e0a\u65e5\u5dee\u5f02 row uses a different column per currency.",
         "These are reproduced exactly so the figures tie back to the existing report. "
         "Set QUIRKS_ASLI = False in mtoatd_spec.py to apply one uniform rule instead."),
    ]
    for name, meaning, note in mtoatd_rows:
        r += 1
        for c, t in ((2, name), (3, meaning), (4, note)):
            cell = ws.cell(r, c, t)
            cell.border = box
            cell.alignment = wrap
            if c == 2:
                cell.font = Font(bold=True)
        ws.row_dimensions[r].height = 48

    r += 2
    ws.cell(r, 2, "Tip: hover over any of the four column headers in sheet 'D' or 'W' "
                  "to see the same explanation as a cell comment.").font = Font(italic=True, size=9, color="808080")
    return ws


def main():
    ap = argparse.ArgumentParser(description="Auto-hitung Handling Fee / Xero Rate / Xero USD / Forex Gain-Loss")
    ap.add_argument("input", help="file .xlsx sumber")
    ap.add_argument("-o", "--output", help="file .xlsx hasil (default: <input>-hasil.xlsx)")
    ap.add_argument("--in-place", action="store_true", help="timpa file input (file harus tertutup di Excel)")
    ap.add_argument("--sheet", nargs="+", help=f"nama sheet data (default: cari {DATA_SHEET_NAMES})")
    ap.add_argument("--fee-from", help=f"ambil sheet '{FEE_SHEET}' dari workbook lain")
    ap.add_argument("--xero-from", help=f"ambil sheet '{XERO_SHEET}' dari workbook lain")
    ap.add_argument("--no-channel", action="store_true",
                    help=f"jangan tulis sheet '{PCB_SHEET}' dan '{JW_SHEET}'")
    ap.add_argument("--no-mtoatd", action="store_true",
                    help=f"jangan tulis sheet '{MTOATD_SHEET}'")
    ap.add_argument("--no-fee-sheet", action="store_true",
                    help=f"jangan tulis sheet '{FEE_SHEET}' hasil gabungan")
    ap.add_argument("--no-report", action="store_true",
                    help=f"jangan buat sheet '{REPORT_SHEET}', '{DETAIL_SHEET}', "
                         f"'{MISSING_SHEET}'")
    ap.add_argument("--open", action="store_true", dest="buka",
                    help="buka file hasil setelah selesai")
    ap.add_argument("--jwallet-opening", metavar="ANGKA",
                    help="saldo penutup J Wallet bulan SEBELUMNYA, dipakai sebagai "
                         "saldo pembuka. Menang atas sheet 'J Wallet'/'Opening Balance'. "
                         "Tanpa ini saldo mulai NOL kalau file tidak punya sheet saldo.")
    ap.add_argument("--period", metavar="YYYY-MM",
                    help="BULAN LAPORAN. Baris D/W dan Fund Transfer Table di luar "
                         "bulan ini DIBUANG (ekspor back office selalu punya ekor "
                         "bulan sebelum/sesudah). Tanpa ini semua tanggal dihitung "
                         "dan periode di sheet report = bulan terbanyak.")
    args = ap.parse_args()

    # Dipasang PALING AWAL: process_sheet dan baca_fund_transfer membacanya lewat
    # global PERIODE_FILTER.
    set_periode(args.period)

    global SALDO_AWAL_JW
    SALDO_AWAL_JW = baca_angka_saldo(args.jwallet_opening)
    if SALDO_AWAL_JW is not None:
        tgl_ob = tanggal_saldo_pembuka()
        print(f"Saldo  : J Wallet mulai dari {SALDO_AWAL_JW:,.2f}"
              + (f" per {tgl_ob}" if tgl_ob else "") + " (diketik di halaman upload)")

    global MAKE_REPORT_SHEET, WRITE_FEE_SHEET, MAKE_MTOATD_SHEET, MAKE_CHANNEL_SHEETS
    if args.no_channel:
        MAKE_CHANNEL_SHEETS = False
    if args.no_mtoatd:
        MAKE_MTOATD_SHEET = False
    if args.no_report:
        MAKE_REPORT_SHEET = False
    if args.no_fee_sheet:
        WRITE_FEE_SHEET = False

    src = Path(args.input).expanduser()
    if not src.exists():
        sys.exit(f"File tidak ditemukan: {src}")
    dst = src if args.in_place else Path(args.output).expanduser() if args.output \
        else src.with_name(f"{src.stem}-hasil.xlsx")

    ukuran = src.stat().st_size / 1024 / 1024
    print(f"Baca   : {src}  ({ukuran:,.1f} MB)")
    if ukuran > 20:
        print("         file besar -- membuka bisa 1-2 menit, mohon tunggu")
    with Denyut(f"membuka {src.name} ({ukuran:,.0f} MB)"):
        wb = openpyxl.load_workbook(src)
    lapor("workbook terbuka")

    def ref_wb(path, sheet, flag):
        """Workbook sumber untuk sheet acuan (fee/kurs)."""
        if path:
            f = Path(path).expanduser()
            if not f.exists():
                sys.exit(f"File acuan tidak ditemukan: {f}")
            print(f"Acuan  : sheet '{sheet}' dari {f.name}")
            book = openpyxl.load_workbook(f, data_only=True)
        else:
            book = wb
        if sheet == XERO_SHEET:
            if cari_sheet(book, XERO_SHEET_NAMA) is None:
                sys.exit(f"Sheet kurs (nama: {' / '.join(XERO_SHEET_NAMA)}) tidak ada di "
                         f"{'file acuan' if path else src.name}.\n"
                         f"Sheet yang ada: {book.sheetnames}\n"
                         f"Tip: pakai {flag} \"file-yang-punya-sheet-itu.xlsx\"")
        elif not cari_sheet_fee(book):
            sys.exit(f"Sheet fee (nama dimulai: {' / '.join(FEE_SHEET_AWALAN)}) tidak ada di "
                     f"{'file acuan' if path else src.name}.\n"
                     f"Sheet yang ada: {book.sheetnames}\n"
                     f"Tip: pakai {flag} \"file-yang-punya-sheet-itu.xlsx\"")
        return book

    fee_table = load_fee_table(ref_wb(args.fee_from, FEE_SHEET, "--fee-from"))
    rates, xero_dates = load_xero(ref_wb(args.xero_from, XERO_SHEET, "--xero-from"))

    if not fee_table:
        sys.exit(f"Sheet '{FEE_SHEET}' tidak berisi satu pun rate.\n"
                 f"Header 'Currency | Payment Gateway | Deposit | Withdrawal' harus di baris "
                 f"{FEE_HEADER_ROW}, datanya mulai baris {FEE_HEADER_ROW + 1}.")
    if not xero_dates:
        sys.exit(f"Sheet '{XERO_SHEET}' tidak berisi satu pun tanggal/kurs.\n"
                 "Baris 1 harus berisi 'Date' lalu kode currency (AED, INR, VND, ...),\n"
                 "baris berikutnya tanggal + kursnya.\n"
                 f"Kalau kurs ada di file lain, pakai --xero-from \"file.xlsx\"")

    print(f"Fee    : {len(fee_table)} kombinasi currency+gateway")
    print(f"XERO   : {len(xero_dates)} tanggal ({xero_dates[0]} s/d {xero_dates[-1]})")
    if PERIODE_FILTER:
        print(f"Periode: {nama_periode()} -- HANYA bulan ini yang dihitung; "
              f"baris di luar bulan ini dibuang")
    else:
        print("Periode: SEMUA TANGGAL (tidak ada --period) -- laporan bisa memuat "
              "ekor bulan lain")

    if args.sheet:
        targets = args.sheet
    else:
        lewati = {norm(XERO_SHEET), norm(LEGEND_SHEET), norm(REPORT_SHEET),
                  norm(DETAIL_SHEET), norm(MISSING_SHEET)}
        lewati |= {norm(x) for x in cari_sheet_fee(wb)}
        # 1. sheet dgn nama yang dikenali, dicoba dari kelompok paling spesifik
        targets = []
        for grup in DATA_SHEET_GROUPS:
            targets = [s for s in wb.sheetnames if norm(s) in grup]
            if targets:
                break
        if targets:
            # Sudah ketemu sheet bernama jelas -> HANYA itu yang dihitung.
            # Workbook besar sering punya banyak sheet transaksi lain
            # ('Deposits', 'Withdrawal (2)', 'Rebate Withdrawal', ...) yang
            # bukan input perhitungan. Pakai --sheet untuk menambahkannya.
            lain = [x for x in wb.sheetnames
                    if x not in targets and norm(x) not in lewati and is_data_sheet(wb[x])]
            if lain:
                print(f"Lewati : {len(lain)} sheet transaksi lain tidak dihitung "
                      f"({', '.join(repr(x) for x in lain[:6])}"
                      f"{', ...' if len(lain) > 6 else ''})")
                print(f"         tambahkan dengan --sheet \"NAMA SHEET\" kalau perlu")
        else:
            # 2. tidak ada nama yang dikenali -> deteksi dari header
            for s_ in wb.sheetnames:
                if norm(s_) in lewati:
                    continue
                if is_data_sheet(wb[s_]):
                    targets.append(s_)
                    print(f"Deteksi: sheet '{s_}' berisi kolom transaksi -> ikut dihitung")
    if not targets:
        sys.exit("Sheet data tidak ketemu.\n"
                 f"Sheet yang ada: {wb.sheetnames}\n"
                 f"Sheet transaksi harus punya kolom: {', '.join(REQUIRED_COLUMNS)}\n"
                 f"  + salah satu kolom tanggal: {' / '.join(DATE_COLUMNS)}\n"
                 "Atau tentukan manual dengan --sheet \"NAMA SHEET\"")

    report = {"fee_match": Counter(), "fee_missing": Counter(), "rate_missing": Counter(),
              "rate_fallback": Counter(), "total_fee": defaultdict(float), "total_forex": 0.0,
              "total_xero_usd": 0.0, "total_usd": 0.0, "errors": [],
              "per_sheet": defaultdict(lambda: {"n": 0, "fee": defaultdict(float),
                                                "forex": 0.0, "xero_usd": 0.0, "usd": 0.0}),
              "records": [], "fx_detail": {}, "data_issues": Counter(), "pivot": {},
              "ditolak": Counter(), "input_kurang": [],
              "dibuang": Counter(), "dibuang_usd": defaultdict(float),
              "luar_periode": Counter(), "luar_periode_usd": defaultdict(float),
              "refuse_mtoatd": []}

    total_rows = 0
    date_cols_used = {}
    for name in targets:
        real = next((s for s in wb.sheetnames if norm(s) == norm(name)), name)
        n, date_col = process_sheet(wb[real], fee_table, rates, xero_dates, report)
        tgl = f"  (kurs pakai kolom '{date_col.title()}')" if date_col else ""
        print(f"Sheet  : '{real}' -> {n} baris dihitung{tgl}")
        total_rows += n
        if date_col:
            date_cols_used[real] = date_col.title()

    if MAKE_CHANNEL_SHEETS and report["records"]:
        nama_ftt, ftt = baca_fund_transfer(wb, rates, xero_dates, report)
        if nama_ftt is None or not ftt:
            # Pesan untuk Terminal (bahasa Indonesia) dan untuk sheet Excel (English)
            sebab_id = ("sheet 'Fund Transfer Table' tidak ada di file yang diupload"
                        if nama_ftt is None else
                        f"sheet {nama_ftt!r} ada tapi KOSONG (tidak ada baris data)")
            sebab_en = ("the uploaded file has no 'Fund Transfer Table' sheet"
                        if nama_ftt is None else
                        f"sheet {nama_ftt!r} exists but is EMPTY - no data rows")
            print(f"Lewati : {sebab_id} -> kolom 'Fund transfer' di '{PCB_SHEET}' "
                  f"dan seluruh '{JW_SHEET}' akan kosong")
            report["input_kurang"].append((
                "Fund Transfer Table", sebab_en,
                f"'{JW_SHEET}' is empty and the 'Fund transfer' / 'Fund transfer charges' "
                f"columns of '{PCB_SHEET}' are 0.",
                "Paste the fund-transfer rows into the 'Fund Transfer Table' sheet of the "
                "template and upload again. This data is NOT in sheet D or W - it is a "
                "separate export, so it can never be derived from the transactions."))
            ftt = []
        else:
            print(f"Baca   : sheet {nama_ftt!r} -> {len(ftt)} pemindahan dana")
        n_pcb, n_jw = write_channel_sheets(wb, report, ftt)
        if n_pcb:
            print(f"Sheet  : '{PCB_SHEET}' -> {n_pcb:,} baris (Date x Currency x Channel); "
                  f"kolom 'Withdrawal charges' KUNING = boleh disesuaikan manual oleh tim "
                  f"(rumusnya sendiri sudah dikonfirmasi benar 3 Sep 2026)")
            print(f"Sheet  : '{JW_SHEET}' -> {n_jw:,} baris dari Fund Transfer Table")

    if MAKE_MTOATD_SHEET and report["records"]:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        with Denyut(f"menyusun sheet '{MTOATD_SHEET}'"):
            n_hari, n_cur = write_mtoatd_sheet(wb, report)
        if n_hari:
            print(f"Sheet  : '{MTOATD_SHEET}' -> {n_hari} hari x {n_cur} currency, "
                  f"{report.get('mtoatd_sel', 0):,} sel dihitung dari D + W "
                  f"(baris tanpa rumus dibiarkan kosong)")

    if WRITE_FEE_SHEET:
        n_fee = write_fee_sheet(wb, fee_table)
        print(f"Sheet  : '{FEE_SHEET}' -> {n_fee} kombinasi (tabel fee tergabung)")

    if MAKE_REPORT_SHEET:
        # periode report = bulan yang paling banyak muncul di data
        bulan = Counter()
        for rec in report["records"]:
            d = as_date(rec.get("Date"))
            if d:
                bulan[(d.year, d.month)] += 1
        if PERIODE_FILTER:
            periode = datetime.date(PERIODE_FILTER[0], PERIODE_FILTER[1], 1)
        elif bulan:
            th, bl = bulan.most_common(1)[0][0]
            periode = datetime.date(th, bl, 1)
        else:
            periode = None
        # label kolom tanggal per jenis (deposit biasanya 'Paid Date')
        lbl = defaultdict(Counter)
        for rec in report["records"]:
            lbl[rec["Type"]][rec["Date Source"]] += 1
        report["date_col_label"] = {
            k: v.most_common(1)[0][0] for k, v in lbl.items()}

        with Denyut(f"menyusun sheet '{REPORT_SHEET}'"):
            write_dw_report_sheet(wb, report, periode)
        n_grup = len(report["pivot"])
        n_dep = sum(1 for k in report["pivot"] if k[0] == "deposit")
        print(f"Sheet  : '{REPORT_SHEET}' -> ringkasan {n_grup} grup "
              f"({n_dep} deposit + {n_grup - n_dep} withdrawal), periode "
              f"{periode.strftime('%b-%y') if periode else '-'}")
        with Denyut(f"menulis {len(report['records']):,} baris ke sheet '{DETAIL_SHEET}'"):
            n_rep = write_detail_sheet(wb, report)
        print(f"Sheet  : '{DETAIL_SHEET}' -> {n_rep:,} baris (gabungan semua sheet transaksi)")
        write_missing_sheet(wb, report)
        n_fx = sum(d["rows"] for d in report["fx_detail"].values())
        n_fee = sum(report["fee_missing"].values())
        status = (f"{len(report['fx_detail'])} tanggal kurs kurang / {n_fx:,} baris, "
                  f"{len(report['fee_missing'])} rate fee kurang / {n_fee:,} baris")
        print(f"Sheet  : '{MISSING_SHEET}' -> {status}")
    if not report["records"]:
        print("\n" + "!" * 108)
        print("!! TIDAK ADA SATU PUN BARIS TRANSAKSI TERBACA -- file hasilnya KOSONG.")
        if PERIODE_FILTER and report.get("luar_periode"):
            tot = sum(report["luar_periode"].values())
            bln = sorted({b for _s, b in report["luar_periode"]})
            print(f"!! SEBABNYA BULAN LAPORAN: kamu memilih {nama_periode()}, tapi "
                  f"{tot:,} baris di file")
            print(f"!! itu tanggalnya {', '.join(bln)} -- tidak ada satu pun yang "
                  f"jatuh di {nama_periode()}.")
            print("!! Pilih bulan yang benar, atau upload file yang memuat bulan itu.")
        else:
            print("!! Sheet 'D' dan 'W' di file yang kamu upload tidak berisi baris data.")
            print("!! Tempel data deposit ke sheet 'D' mulai baris 2, withdrawal ke 'W' mulai")
            print("!! baris 2, biarkan baris header-nya, lalu jalankan ulang.")
        print("!" * 108)

    write_legend_sheet(wb, date_cols_used)
    print(f"Sheet  : '{LEGEND_SHEET}' -> penjelasan kolom & warna (English)")

    # ---------------- ringkasan ----------------
    print("\n" + "=" * 108)
    print("RATE YANG DIPAKAI")
    print("=" * 108)
    print(f"{'CUR':5} {'GATEWAY (data)':18} {'JENIS':11} {'n':>5} -> {'BARIS FEE':22} "
          f"{'MATCH':14} {'RATE':>8} {'+TETAP':>8} {'MIN':>7} {'LAMA':>7}")
    for (cur, gw, key, how, pct, tetap, minim, legacy, kind), n in sorted(
            report["fee_match"].items(), key=lambda x: -x[1]):
        flag = "  <-- CEK!" if how == "fallback-other" else ""
        print(f"{cur:5} {gw:18} {kind:11} {n:>5} -> {key:22} {how:14} {pct*100:7.4f}% "
              f"{tetap:>8,.0f} {('-' if minim is None else f'{minim:,.0f}'):>7} "
              f"{legacy:>7,.0f}{flag}")

    if report.get("luar_periode"):
        tot = sum(report["luar_periode"].values())
        print(f"\n!  DI LUAR {nama_periode().upper()}: {tot:,} baris dibuang "
              f"(tanggalnya bukan bulan laporan)")
        for (sh, bln), jml in sorted(report["luar_periode"].items()):
            usd = report["luar_periode_usd"].get((sh, bln), 0.0)
            ket = f"  USD {usd:>12,.2f}" if usd else ""
            print(f"   sheet {sh:20} {bln:10} {jml:>7,} baris{ket}")

    if report.get("dibuang"):
        tot = sum(report["dibuang"].values())
        print(f"\n!  DIBUANG: {tot:,} baris statusnya BUKAN 'finish' -> tidak dihitung "
              f"di laporan mana pun")
        for (sh, cur, st), n in sorted(report["dibuang"].items(), key=lambda x: -x[1])[:15]:
            usd = report["dibuang_usd"][(sh, cur, st)]
            print(f"   sheet {sh:3} {cur:6} status {st.lower():10} {n:>5} baris  "
                  f"USD {usd:>12,.2f}")

    if report["fee_missing"]:
        tot = sum(report["fee_missing"].values())
        print(f"\n!! BUTUH DATA RATE: {tot} baris -> Handling Fee dikosongkan & ditandai MERAH di Excel")
        for (cur, gw), n in sorted(report["fee_missing"].items(), key=lambda x: -x[1]):
            print(f"   {cur:6} / {gw:22} {n:>5} baris")

    if report["rate_fallback"]:
        tot = sum(report["rate_fallback"].values())
        print(f"\n!  KURS FALLBACK: {tot} baris Paid Date-nya di luar sheet XERO -> pakai tanggal terdekat")
        print("   (sel 'Xero Rate'-nya ditandai kuning di Excel)")
        for (d, cur, used), n in sorted(report["rate_fallback"].items())[:12]:
            print(f"   {d} {cur:5} -> pakai kurs {used}  ({n} baris)")
        if len(report["rate_fallback"]) > 12:
            print(f"   ... dan {len(report['rate_fallback']) - 12} kombinasi lain")

    if report["rate_missing"]:
        print("\n!! KURS TIDAK KETEMU (Xero Rate/USD/Forex dikosongkan):")
        for (d, cur), n in list(report["rate_missing"].items())[:12]:
            print(f"   {d} {cur}  ({n} baris)")

    print("\n" + "=" * 108)
    print("TOTAL")
    print("=" * 108)
    if len(report["per_sheet"]) > 1:
        for name, sub in report["per_sheet"].items():
            print(f"  -- sheet '{name}' ({sub['n']:,} baris)")
            for cur, v in sorted(sub["fee"].items()):
                print(f"     Handling Fee {cur:6}: {v:>20,.2f}")
            print(f"     {'USD (kolom lama)':19}: {sub['usd']:>20,.2f}")
            print(f"     {'Xero USD':19}: {sub['xero_usd']:>20,.2f}")
            print(f"     {'Forex Gain/Loss':19}: {sub['forex']:>20,.2f}")
        print("  -- GABUNGAN")
    for cur, v in sorted(report["total_fee"].items()):
        print(f"  Handling Fee {cur:6} : {v:>20,.2f}")
    print(f"  {'USD (kolom lama)':21}: {report['total_usd']:>20,.2f}")
    print(f"  {'Xero USD':21}: {report['total_xero_usd']:>20,.2f}")
    print(f"  {'Forex Gain/Loss':21}: {report['total_forex']:>20,.2f}")
    print(f"  {'Baris dihitung':21}: {total_rows:>20,}")

    for e in report["errors"]:
        print("ERROR:", e)

    # --- urutkan tab sesuai URUTAN_SHEET
    if URUTAN_SHEET:
        depan = [n for x in URUTAN_SHEET
                 for n in wb.sheetnames if norm(n) == norm(x)]
        sisa = [n for n in wb.sheetnames if n not in depan]
        wb._sheets = [wb[n] for n in depan + sisa]
        print(f"Urutan : {' | '.join(wb.sheetnames)}")

    try:
        with Denyut(f"menyimpan {dst.name}"):
            wb.save(dst)
    except PermissionError:
        sys.exit(f"\nGagal simpan: '{dst}' sedang dibuka di Excel. Tutup dulu, lalu jalankan ulang.")
    print(f"\nSimpan : {dst}")
    print(f"Selesai dalam {time.monotonic() - _T0:.0f} detik")

    if args.buka:
        import subprocess
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", str(dst)], check=False)
            elif sys.platform.startswith("win"):
                os.startfile(str(dst))          # noqa: S606
            else:
                subprocess.run(["xdg-open", str(dst)], check=False)
        except Exception as e:
            print(f"(tidak bisa membuka otomatis: {e})")


if __name__ == "__main__":
    main()
