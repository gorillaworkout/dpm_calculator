#!/usr/bin/env python3
"""Spesifikasi + mesin hitung sheet MTOATD.

MTOATD adalah rekonsiliasi HARIAN antara empat sistem:
    MT4     platform trading
    Wallet  dompet
    CRM     back office
    TD      data pembayaran (inilah sheet 'D' dan 'W')

Satu blok = satu tanggal = 51 baris. Kolom: label lalu satu kolom per currency,
ditutup kolom 合計 (jumlah antar currency), lalu blok kedua 当月累计 (akumulasi
sejak awal bulan) dengan susunan baris yang sama.

SEMUA baris yang bisa dihitung, dihitung dari sheet D + W saja
=============================================================
Sebelumnya baris MT4 / Wallet / CRM dianggap harus diisi manusia. Itu KELIRU.
Workbook sumber `22.08 Bayu - D&W-Dupoin Markets-JUN 2026v3.xlsx` ternyata
menyimpan rumusnya di sheet 'MTOATD (MAY''26)' -- semuanya SUMIFS ke sheet
'Deposits' / 'Withdrawals', yang isi & susunan kolomnya identik dengan
sheet 'D' / 'W'. Jadi tidak ada input manual lagi: sekali upload, jadi.

Peta kolom (sama untuk D<->Deposits dan W<->Withdrawals):
    Deposits    H Paid Date   J Settlement Date  K Currency  N USD  P Transaction
    Withdrawals J Currency  M USD  O Charges  P Transaction
                S Completed Date  U Settlement Date  V Status  A Source Name

Rumus per tanggal t, per currency:
    MT4入金            = Σ USD          [Settlement Date = t]
    Wallet入金          = (tidak ada rumus -> dibiarkan KOSONG)
    CRM入金            = MT4入金 + Wallet入金
    TD入金             = Σ USD          [Paid Date = t]
    ——crm已入,TD未入    = Σ USD          [Settlement = t, Paid > t]
    ——调上日差异 (入金)   = -Σ USD         [Paid = t, Settlement < t]
    MT4出金            = Σ USD          [Settlement = t, Source Name = 'Withdrawal']
    Wallet出金          = Σ USD          [Settlement = t, Source Name = 'Rebate Withdrawal']
    CRM出金            = MT4出金 + Wallet出金
    TD出金             = Σ USD          [Completed = t, Status ≠ refuse]
    ——crm已出,TD未出    = Σ USD          [Settlement = t, Completed > t, Status ≠ refuse]
                        + Σ USD          [Settlement = t, Completed = 1/1/1970]
    ——调上日差异 (出金)   = -Σ USD         [Completed = t, Settlement < t, Status ≠ refuse]
    CRM 入金（原币种）   = Σ Transaction  [Settlement = t]
    CRM 出金（原币种）   = Σ OC           [Settlement = t]
    实收（原币种）       = Σ Transaction  [Paid = t]
    实出（原币种）       = Σ Transaction  [Completed = t]
    ——未打款            = Σ OC           [Settlement = t, Completed > t, Status ≠ refuse]
                        + Σ OC           [Settlement = t, Completed = 1/1/1970]
    ——手续费/转账费      = Σ Charges      [Completed = t]      (hanya USDT & USD)
    ——调上日差异 (原币种) = -Σ Transaction [Completed = t, Settlement < t]
    selisih             = pengurangan antar baris di atas (lihat TURUNAN)
    本日 MT4+錢包淨入金  = MT4入金 + Wallet入金 - MT4出金 - Wallet出金
    月累计              = akumulasi 本日 sepanjang bulan
    合計                = jumlah antar currency

OC = "original currency": untuk USDT dan USD mata uang aslinya memang USD,
jadi rumus mereka memakai kolom USD; currency lain memakai kolom Transaction.

Baris yang di workbook sumber TIDAK punya rumus (rincian manual seperti
——作废单, ——小数点差异, Wallet入金) dibiarkan KOSONG, bukan diisi nol.

Terverifikasi 523/523 sel terhadap 1 Jun 2026 (19 currency + kolom 合計).
"""

import datetime

EPOCH = datetime.date(1970, 1, 1)          # CRM menandai 'belum dibayar' begini

# Currency yang baris （原币种）-nya membaca kolom USD (M), bukan kolom
# Transaction (P).
#
# KOSONG = baris （原币种）SELALU kolom P (Transaction), untuk SEMUA currency.
#
# Riwayat bolak-balik, jangan diulang:
#   3 Sep 2026  diubah jadi kolom P, lalu DIKEMBALIKAN ke M karena tim menjawab
#               "keep the existing formula, USDT & USD only has 'USD' to capture".
#   7 Sep 2026  tim mengirim ANGKA ACUAN yang membantah jawaban mereka sendiri:
#               CRM出金（原币种）USDT 17 Jun = 73.561,09 + 30.420,82 = 103.981,91.
#               Diukur dari data: 73.561,09 itu kolom P (kolom M = 73.658,09) dan
#               30.420,82 juga kolom P (kolom M = 30.425,32).
#               Jadi kolom P yang BENAR. Angka acuan menang atas jawaban lisan.
OC_KOLOM_USD = ()

# Currency yang PUNYA baris ——手续费/转账费 di sheet mereka (cuma dua ini; currency
# lain barisnya memang kosong). Ini BEDA URUSAN dengan OC_KOLOM_USD di atas --
# dulu keduanya memakai satu konstanta yang sama, jadi mengosongkan yang satu
# ikut menghapus baris fee. Sekarang dipisah.
OC_ADA_BARIS_FEE = ("USDT", "USD")

# Urutan kolom currency, mengikuti sheet MTOATD yang sudah beredar.
# Currency lain yang muncul di data ditambahkan di belakang.
URUT_CURRENCY = ("VND", "INR", "JPY", "LAK", "THB", "TRY", "PKR", "EGP", "KES",
                 "KRW", "AED", "BRL", "MXN", "NGN", "ZAR", "PHP", "KHR",
                 "USDT", "USD")

# Ikut meniru kejanggalan rumus asli? (lihat catatan di hitung() di bawah)
# SUDAH TIDAK DIPAKAI sejak 7 Sep 2026 -- dibiarkan supaya file lama yang meng-
# import-nya tidak error. Semua kejanggalan khusus kolom USDT/USD sudah dihapus:
# baris （原币种）sekarang SERAGAM kolom P untuk semua currency, dan semua baris
# berbasis Completed Date seragam membuang refuse. Lihat OC_KOLOM_USD di atas.
QUIRKS_ASLI = True   # tidak lagi mengubah apa pun

# (offset, label, kategori kolom A, jenis, kunci)
#   header  baris tanggal
#   calc    dihitung dari sheet D / W
#   manual  di workbook sumber tidak ada rumusnya -> dibiarkan kosong
#   blank   baris pemisah
BARIS = [
    (0,  None,                              "日期：", "header", None),
    (1,  "MT4入金 (USD)",                    "入金",  "calc",   "mt4_d"),
    (2,  "Wallet入金 (USD)",                 None,   "manual", "wallet_d"),
    (3,  "CRM入金 (USD)",                    None,   "calc",   "crm_d"),
    (4,  "TD入金 (USD)",                     None,   "calc",   "td_d"),
    (5,  "MT4+Wallet-CRM 入金差异 (USD)",     None,   "calc",   "var_d_mwc"),
    (6,  "——mt4已入,crm未入",                 None,   "manual", None),
    (7,  "——mt4未入,crm已入",                 None,   "manual", None),
    (8,  "——调上日差异",                      None,   "manual", None),
    (9,  "——小数点合计差异",                    None,   "manual", None),
    (10, "CRM-TD 入金差异 (USD)",             None,   "calc",   "var_d_ct"),
    (11, "——crm已入,TD未入",                  None,   "calc",   "d_crm_in_td_out"),
    (12, "——crm未入,TD已入",                  None,   "manual", None),
    (13, "——调上日差异",                      None,   "calc",   "d_adj_prev"),
    (14, "——小数点合计差异",                    None,   "manual", None),
    (15, "MT4出金 (USD)",                    "出金",  "calc",   "mt4_w"),
    (16, "Wallet出金 (USD)",                 None,   "calc",   "wallet_w"),
    (17, "CRM出金 (USD)",                    None,   "calc",   "crm_w"),
    (18, "TD出金 (USD)",                     None,   "calc",   "td_w"),
    (19, "MT4+Wallet-CRM   出金差异 (USD)",   None,   "calc",   "var_w_mwc"),
    (20, "——作废单",                         None,   "manual", None),
    (21, "——mt4已出,crm未出",                 None,   "manual", None),
    (22, "——mt4未出,crm已出",                 None,   "manual", None),
    (23, "——调上日差异",                      None,   "manual", None),
    (24, "CRM-TD  出金差异 (USD)",            None,   "calc",   "var_w_ct"),
    (25, "——作废单",                         None,   "manual", None),
    (26, "——crm已出,TD未出",                  None,   "calc",   "w_crm_out_td_out"),
    (27, "——crm未出,TD已出",                  None,   "manual", None),
    (28, "——调上日差异",                      None,   "calc",   "w_adj_prev"),
    (29, None,                              None,   "blank",  None),
    (30, "CRM 入金 （原币种）",                None,   "calc",   "crm_d_oc"),
    (31, "CRM 出金 （原币种）",                None,   "calc",   "crm_w_oc"),
    (32, "实收（原币种）",                     None,   "calc",   "recv_oc"),
    (33, "实出（原币种）",                     None,   "calc",   "pay_oc"),
    (34, "CRM入金 -实收差异",                  None,   "calc",   "var_oc_d"),
    (35, "——crm已入,TD未入",                  None,   "calc",   "oc_crm_in_td_out"),
    (36, "——crm未入,TD已入",                  None,   "manual", None),
    (37, "——调上日差异",                      None,   "calc",   "oc_d_adj_prev"),
    (38, "CRM出金 -实出差异",                  None,   "calc",   "var_oc_w"),
    (39, "——作废单",                         None,   "manual", None),
    (40, "——未打款",                         None,   "calc",   "oc_belum_bayar"),
    (41, "——手续费/转账费",                    None,   "calc",   "oc_charges"),
    (42, "——MYR改USDT出",                    None,   "manual", None),
    (43, "——小数点差异",                      None,   "manual", None),
    (44, "——分拆出金",                        None,   "manual", None),
    (45, "——调上日差异",                      None,   "calc",   "oc_w_adj_prev"),
    (46, "本 日   MT4+錢包淨入金 (USD)",       None,   "calc",   "net_hari"),
    (47, "月累计 MT4+錢包淨入金 (USD)",         None,   "cum",    "net_bulan"),
    (48, None,                              None,   "blank",  None),
    (49, None,                              None,   "blank",  None),
    (50, None,                              None,   "blank",  None),
]
TINGGI_BLOK = 51
KOL_LABEL_A, KOL_LABEL_B, KOL_CUR_AWAL = 1, 2, 3   # A=kategori, B=label, C..=currency

OFFSET_NET, OFFSET_CUM = 46, 47

# Terjemahan Inggris untuk komentar di sel label
EN = {
    "MT4入金 (USD)": "MT4 Deposit (USD) = sum of USD where Settlement Date = this day",
    "Wallet入金 (USD)": "Wallet Deposit (USD) - no formula in the source workbook, left blank",
    "CRM入金 (USD)": "CRM Deposit (USD) = MT4 Deposit + Wallet Deposit",
    "TD入金 (USD)": "TD Deposit (USD) = sum of USD where Paid Date = this day",
    "MT4+Wallet-CRM 入金差异 (USD)": "MT4 + Wallet - CRM Deposit Variance",
    "CRM-TD 入金差异 (USD)": "CRM - TD Deposit Variance",
    "MT4出金 (USD)": "MT4 Withdrawal (USD) = sum of USD where Settlement Date = this day "
                     "and Source Name = 'Withdrawal'",
    "Wallet出金 (USD)": "Wallet Withdrawal (USD) = sum of USD where Settlement Date = this day "
                        "and Source Name = 'Rebate Withdrawal'",
    "CRM出金 (USD)": "CRM Withdrawal (USD) = MT4 Withdrawal + Wallet Withdrawal",
    "TD出金 (USD)": "TD Withdrawal (USD) = sum of USD where Completed Date = this day "
                    "and Status is not 'refuse'",
    "MT4+Wallet-CRM   出金差异 (USD)": "MT4 + Wallet - CRM Withdrawal Variance",
    "CRM-TD  出金差异 (USD)": "CRM - TD Withdrawal Variance",
    "CRM 入金 （原币种）": "CRM Deposit (Original Currency) = sum of Transaction "
                        "where Settlement Date = this day",
    "CRM 出金 （原币种）": "CRM Withdrawal (Original Currency) = sum of Transaction "
                        "where Settlement Date = this day (USDT/USD use the USD column)",
    "实收（原币种）": "Actual Receipt (Original Currency) = sum of Transaction "
                    "where Paid Date = this day",
    "实出（原币种）": "Actual Payout (Original Currency) = sum of Transaction "
                    "where Completed Date = this day",
    "CRM入金 -实收差异": "CRM Deposit vs Actual Receipt Variance",
    "CRM出金 -实出差异": "CRM Withdrawal vs Actual Payout Variance",
    "本 日   MT4+錢包淨入金 (USD)": "Current day MT4 + Wallet Net Deposit (USD)",
    "月累计 MT4+錢包淨入金 (USD)": "Month-to-date MT4 + Wallet Net Deposit (USD)",
    "——mt4已入,crm未入": "MT4 deposited, not yet in CRM - manual breakdown, no formula",
    "——mt4未入,crm已入": "not yet in MT4, already in CRM - manual breakdown, no formula",
    "——crm已入,TD未入": "CRM deposited, not yet in TD = Settlement Date is this day "
                       "but Paid Date is later",
    "——crm未入,TD已入": "not yet in CRM, already in TD - manual breakdown, no formula",
    "——mt4已出,crm未出": "MT4 withdrawn, not yet in CRM - manual breakdown, no formula",
    "——mt4未出,crm已出": "not yet in MT4, already in CRM - manual breakdown, no formula",
    "——crm已出,TD未出": "CRM withdrawn, not yet in TD = Settlement Date is this day but "
                       "Completed Date is later (or 1/1/1970 = not paid yet)",
    "——crm未出,TD已出": "not yet in CRM, already in TD - manual breakdown, no formula",
    "——调上日差异": "previous day's variance adjustment = booked on this day but "
                  "settled on an earlier day",
    "——小数点合计差异": "small-amount rounding variance - manual breakdown, no formula",
    "——小数点差异": "rounding variance - manual breakdown, no formula",
    "——作废单": "voided / pending-review order - manual breakdown, no formula",
    "——未打款": "pending payment = Settlement Date is this day but not paid out yet",
    "——手续费/转账费": "handling fee / transfer fee = sum of the Charges column "
                     "(USDT and USD only, same as the source workbook)",
    "——MYR改USDT出": "MYR converted to USDT withdrawal - manual breakdown, no formula",
    "——分拆出金": "split withdrawal - manual breakdown, no formula",
}

KUNCI_CALC = [k for _o, _l, _kat, j, k in BARIS if j in ("calc", "cum") and k]
MANUAL_OFFSETS = [o for o, _l, _kat, j, _k in BARIS if j == "manual"]


# ---------------------------------------------------------------------------
# mesin hitung
# ---------------------------------------------------------------------------
class Indeks:
    """Baris D & W dikelompokkan per currency lalu per tanggal, supaya
    perhitungan 31 hari x 20 currency tidak menyapu seluruh data tiap kali.

    Baris deposit  : (paid, settle, usd, txn)
    Baris withdraw : (completed, settle, usd, txn, charges, status, source)
    """

    def __init__(self):
        self.dep_sd, self.dep_pd = {}, {}      # (cur, tgl) -> list baris
        self.wd_sw, self.wd_cd = {}, {}
        self.currencies = set()
        # tanggal blok MTOATD = tanggal transaksinya (Paid Date / Completed Date),
        # BUKAN Settlement Date. Kalau Settlement Date ikut, satu blok tanggal
        # bulan sebelumnya akan muncul (di Juni 2026 ada settlement 31 Mei).
        self.tanggal = set()

    def tambah_deposit(self, cur, paid, settle, usd, txn):
        row = (paid, settle, usd or 0.0, txn or 0.0)
        self.currencies.add(cur)
        if settle:
            self.dep_sd.setdefault((cur, settle), []).append(row)
        if paid:
            self.dep_pd.setdefault((cur, paid), []).append(row)
            self.tanggal.add(paid)

    def tambah_withdrawal(self, cur, completed, settle, usd, txn, charges, status, source,
                          daftar_tanggal=True):
        """daftar_tanggal=False -> baris ikut dijumlah tapi TIDAK membuat blok
        tanggal baru. Dipakai untuk baris `refuse` yang hanya diperlukan oleh
        aturan H+1 di baris berbasis Settlement Date."""
        row = (completed, settle, usd or 0.0, txn or 0.0, charges or 0.0,
               (status or "").strip().lower(), (source or "").strip())
        self.currencies.add(cur)
        if settle:
            self.wd_sw.setdefault((cur, settle), []).append(row)
        if completed and completed != EPOCH:
            self.wd_cd.setdefault((cur, completed), []).append(row)
            if daftar_tanggal:
                self.tanggal.add(completed)


def hitung(idx, cur, t):
    """Kembalikan {kunci: nilai} untuk satu currency di satu tanggal.

    (RIWAYAT) Dulu tiga kejanggalan rumus asli ditiru. SUDAH TIDAK LAGI sejak
    7 Sep 2026 -- semuanya seragam sekarang. Teks di bawah disimpan sebagai catatan
    kenapa rumus asli mereka terlihat aneh. Dulu berbunyi: kejanggalan ditiru kalau
    QUIRKS_ASLI = True. Semuanya
    hasil rumus yang di-drag lalu diedit sebagian di kolom USDT / USD saja:

      实出（原币种）      hanya kolom USDT yang menyaring Status ≠ refuse
      ——调上日差异 (原币种) selalu memakai Transaction dan menyaring refuse
      ——手续费/转账费     hanya ada di kolom USDT & USD; USDT menyaring refuse, USD tidak

    ATURAN YANG BERLAKU SEKARANG (dikonfirmasi tim Malaysia lewat angka acuan,
    7 Sep 2026 -- lihat verifikasi/cek_acuan_tim.py, harus 10/10):
      1. Baris （原币种）SELALU kolom P (Transaction), termasuk USDT & USD.
      2. Baris berbasis SETTLEMENT DATE (mt4_w / wallet_w / crm_w_oc) ikut
         menghitung baris `refuse` yang Completed Date = Settlement Date + 1 hari.
      3. Baris berbasis COMPLETED DATE (td_w / pay_oc / oc_belum_bayar /
         oc_w_adj_prev / oc_charges) TETAP membuang refuse sepenuhnya.
    """
    dep_sd = idx.dep_sd.get((cur, t), ())
    dep_pd = idx.dep_pd.get((cur, t), ())
    wd_sw = idx.wd_sw.get((cur, t), ())
    wd_cd = idx.wd_cd.get((cur, t), ())

    oc = 2 if cur in OC_KOLOM_USD else 3        # indeks kolom OC di baris withdrawal
    S = lambda rows, i: sum(r[i] for r in rows)
    lolos = lambda r: r[5] != "refuse"

    # ATURAN TIM MALAYSIA 7 Sep 2026 -- baris berbasis SETTLEMENT DATE
    # (MT4出金 / Wallet出金 / CRM出金（原币种）) ikut menghitung baris `refuse`
    # yang Completed Date-nya TEPAT SEHARI setelah Settlement Date.
    # Alasannya: uang itu memang keluar dari MT4 pada tanggal settlement dan baru
    # dikembalikan keesokan harinya, jadi pada hari itu dia memang sedang keluar.
    # Terverifikasi: USDT 17 Jun, 4 baris refuse Completed 18 Jun -> MT4出金
    # 42.133,09 + 30.425,32 = 72.558,41 (angka acuan tim).
    # HANYA berlaku di baris Settlement. Baris berbasis Completed Date
    # (TD出金, 实出（原币种）, dst) TETAP membuang refuse sepenuhnya.
    besok = t + datetime.timedelta(days=1)
    ikut_sw = lambda r: r[5] != "refuse" or r[0] == besok

    v = {}
    # ---- 入金 (USD)
    v["mt4_d"] = S(dep_sd, 2)
    v["wallet_d"] = None                                   # tidak ada rumusnya
    v["crm_d"] = v["mt4_d"]
    v["td_d"] = S(dep_pd, 2)
    v["var_d_mwc"] = v["mt4_d"] - v["crm_d"]
    v["var_d_ct"] = v["crm_d"] - v["td_d"]
    v["d_crm_in_td_out"] = S([r for r in dep_sd if r[0] and r[0] > t], 2)
    v["d_adj_prev"] = -S([r for r in dep_pd if r[1] and r[1] < t], 2)

    # ---- 出金 (USD)
    v["mt4_w"] = S([r for r in wd_sw if r[6] == "Withdrawal" and ikut_sw(r)], 2)
    v["wallet_w"] = S([r for r in wd_sw if r[6] == "Rebate Withdrawal" and ikut_sw(r)], 2)
    v["crm_w"] = v["mt4_w"] + v["wallet_w"]
    v["td_w"] = S([r for r in wd_cd if lolos(r)], 2)
    v["var_w_mwc"] = v["mt4_w"] + v["wallet_w"] - v["crm_w"]
    v["var_w_ct"] = v["crm_w"] - v["td_w"]
    v["w_crm_out_td_out"] = (
        S([r for r in wd_sw if r[0] and r[0] > t and lolos(r)], 2)
        + S([r for r in wd_sw if r[0] == EPOCH], 2))
    v["w_adj_prev"] = -S([r for r in wd_cd if r[1] and r[1] < t and lolos(r)], 2)

    # ---- （原币种）
    v["crm_d_oc"] = S(dep_sd, 3)
    v["crm_w_oc"] = S([r for r in wd_sw if ikut_sw(r)], oc)
    v["recv_oc"] = S(dep_pd, 3)
    # 实出（原币种）: di sheet mereka kolom USDT punya saringan $V:$V,"<>refuse"
    # sementara VND/THB/USD tidak. Ditanyakan, dijawab 3 Sep 2026:
    #     "Can remove the '<>refuse', the formula can same as other cells"
    # Jadi seragam, tanpa saringan status. Tidak ada efek ke angka di sini:
    # baris non-finish sudah dibuang lebih dulu di hitung_dw.py (HANYA_STATUS_FINISH),
    # jadi wd_cd memang cuma berisi baris 'finish'.
    v["pay_oc"] = S([r for r in wd_cd if lolos(r)], 3)
    v["var_oc_d"] = v["crm_d_oc"] - v["recv_oc"]
    v["var_oc_w"] = v["crm_w_oc"] - v["pay_oc"]
    v["oc_crm_in_td_out"] = S([r for r in dep_sd if r[0] and r[0] > t], 3)
    # (baris di bawah ini semuanya sudah menyaring refuse lewat lolos())
    v["oc_d_adj_prev"] = -S([r for r in dep_pd if r[1] and r[1] < t], 3)
    v["oc_belum_bayar"] = (
        S([r for r in wd_sw if r[0] and r[0] > t and lolos(r)], oc)
        + S([r for r in wd_sw if r[0] == EPOCH], oc))

    # ——手续费/转账费: di workbook sumber hanya ada di kolom USDT & USD
    if cur in OC_ADA_BARIS_FEE:
        # SELALU buang refuse. Dulu kolom USD tidak menyaring (meniru rumus asli),
        # dan itu tidak masalah selama baris refuse memang tidak pernah masuk index.
        # Sejak 7 Sep 2026 baris refuse IKUT masuk (untuk aturan H+1 di baris
        # Settlement), jadi saringan ini WAJIB -- tanpa itu ——手续费/转账费 kolom USD
        # ikut menghitung charges baris refuse dan keluar dari 570/570.
        v["oc_charges"] = S([r for r in wd_cd if lolos(r)], 4)
    else:
        v["oc_charges"] = None

    # ——调上日差异 (原币种)
    lewat = [r for r in wd_cd if r[1] and r[1] < t and lolos(r)]
    v["oc_w_adj_prev"] = -S(lewat, 3)

    # ---- net harian
    v["net_hari"] = v["mt4_d"] - v["mt4_w"] - v["wallet_w"]
    return v
