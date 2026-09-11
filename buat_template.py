#!/usr/bin/env python3
"""Bikin file template D&W yang bersih dan siap pakai.

Isi template:
    Guide      penjelasan singkat (English) -- apa ditempel di mana
    D          header deposit, kosong, siap ditempel
    W          header withdrawal, kosong, siap ditempel
    Xero       kurs harian, disalin dari file sumber
    D&W FEE    SATU tabel fee -- hasil gabungan 'D&W TD FEE' + 'D&W TD FEE -add.'

Sumber fee & kurs diambil dari workbook yang kamu tunjuk. Logika penggabungannya
sama persis dengan yang dipakai hitung_dw.py, jadi tidak ada selisih.

Pakai:
    python3 buat_template.py "21.08 Bayu - D&W....xlsx"
    python3 buat_template.py "sumber.xlsx" -o "Template D&W.xlsx"
"""

import argparse
import importlib.util
import sys
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl belum terpasang:  python3 -m pip install --user openpyxl")

HDR_D = ["Source Name", "Sales UID", "Sales Name", "Sales Dept", "Sales Area", "User ID",
         "Account", "Paid Date", "Paid Time", "Settlement Date", "Currency", "MT Order",
         "Ticket", "USD", "Rate", "Transaction", "Payment Gateway", "Reference", "Hash",
         "Handling Fee", "Xero Rate", "Xero USD", "Forex Gain/Loss"]
HDR_W = ["Source Name", "Sales UID", "Sales Name", "Sales Dept", "Sales Area", "User ID",
         "Account", "Apply Date", "Apply Time", "Currency", "Payment Gateway", "Reference",
         "USD", "Rate", "Charges", "Transaction", "Trading Charges", "Pay Channel",
         "Completed Date", "Completed Time", "Settlement Date", "Status", "Remark",
         "Handling Fee", "Xero Rate", "Xero USD", "Forex Gain/Loss"]
DIHITUNG = {"Handling Fee", "Xero Rate", "Xero USD", "Forex Gain/Loss"}

# Fund Transfer Table: dua sisi bersebelahan, 汇出 (dana keluar) lalu 收款 (dana masuk).
# Nama kolom sengaja sama dengan sheet aslinya supaya bisa langsung di-copy-paste.
FTT_R2 = {1: "序号", 2: "汇出", 9: "收款", 16: "SPECIFIED DETAILS", 17: "备注"}
FTT_R3 = {2: "汇款日期", 3: "汇出部门", 4: "支付渠道", 5: "汇出币种", 6: "Xero (USD)",
          7: "金额", 8: "手续费", 9: "收款日期", 10: "收款部门", 11: "收款人",
          12: "收入币种", 13: "金额", 14: "手续费", 15: "Xero (USD)", 18: "到账时间"}
FTT_EN = {"汇款日期": "remittance date", "汇出部门": "sending department",
          "支付渠道": "payment channel", "汇出币种": "sending currency",
          "金额": "amount", "手续费": "fee", "收款日期": "receiving date",
          "收款部门": "receiving department", "收款人": "recipient (put 'J Wallet' here "
          "when the money goes into the wallet)", "收入币种": "receiving currency",
          "到账时间": "time credited", "Xero (USD)": "value in USD at the Xero rate"}
WAJIB = {"Currency", "Payment Gateway", "Transaction", "USD",
         "Paid Date", "Completed Date", "Source Name"}

# Saldo pembuka: satu baris per (currency, channel). Baris dompet USDT ditandai
# Channel = 'J Wallet'. Diisi otomatis oleh isi_template.py.
OB_SHEET = "Opening Balance"
OB_KOLOM = ["As Of Date", "Currency", "Channel", "Opening Balance"]
OB_EN = {"As Of Date": "the day this balance is the closing balance of - normally the "
                       "last day of the previous month",
         "Channel": "payment channel name, exactly as it appears in the Payment Gateway "
                    "column. Use 'J Wallet' for the USDT wallet row.",
         "Opening Balance": "in the original currency (the USDT wallet row is in USDT)"}

BIRU, HIJAU, UNGU, ABU = "4472C4", "2E7D32", "7030A0", "808080"
FMT_TGL, FMT_JAM = "yyyy-mm-dd", "hh:mm:ss"


def muat_hitung_dw(folder):
    f = folder / "hitung_dw.py"
    if not f.exists():
        sys.exit(f"hitung_dw.py tidak ada di {folder} -- dibutuhkan untuk logika gabung fee")
    spec = importlib.util.spec_from_file_location("hd", f)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def tulis_header(ws, kolom, warna):
    tipis = Side(style="thin", color="D9D9D9")
    for c, t in enumerate(kolom, start=1):
        sel = ws.cell(1, c, t)
        sel.font = Font(bold=True, color="FFFFFF", size=10)
        sel.fill = PatternFill("solid", fgColor=UNGU if t in DIHITUNG else warna)
        sel.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        sel.border = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)
        lebar = 11
        if t in ("Sales Name", "Sales Dept", "Sales Area", "Payment Gateway", "Pay Channel"):
            lebar = 20
        elif t in ("Hash", "Reference", "Remark"):
            lebar = 18
        elif t in DIHITUNG:
            lebar = 14
        ws.column_dimensions[get_column_letter(c)].width = lebar
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(kolom))}1"


def main():
    ap = argparse.ArgumentParser(description="Bikin template D&W yang siap pakai")
    ap.add_argument("input", help="workbook sumber (untuk tabel fee + kurs Xero)")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()

    src = Path(args.input).expanduser()
    if not src.exists():
        sys.exit(f"File tidak ditemukan: {src}")
    dst = Path(args.output).expanduser() if args.output else src.with_name("Template D&W.xlsx")

    hd = muat_hitung_dw(Path(__file__).resolve().parent)

    print(f"Baca   : {src.name}")
    wb_src = openpyxl.load_workbook(src, read_only=True, data_only=True)
    fee = hd.load_fee_table(wb_src)
    print(f"Fee    : {len(fee)} kombinasi tergabung")

    nama_xero = hd.cari_sheet(wb_src, hd.XERO_SHEET_NAMA)
    kurs = []
    if nama_xero:
        for row in wb_src[nama_xero].iter_rows(values_only=True):
            if any(v not in (None, "") for v in row):
                kurs.append(list(row))
        print(f"Xero   : {len(kurs) - 1} tanggal kurs disalin")
    else:
        print("Xero   : sheet kurs tidak ketemu di sumber -- dibuat kosong")
    wb_src.close()

    out = openpyxl.Workbook()

    # ---------------- Guide ----------------
    ws = out.active
    ws.title = "Guide"
    ws.sheet_properties.tabColor = ABU
    for kol, w in zip("ABC", (3, 26, 96)):
        ws.column_dimensions[kol].width = w
    r = 2
    ws.cell(r, 2, "D&W Calculator - Template").font = Font(bold=True, size=18)
    r += 1
    ws.cell(r, 2, "Paste your data, run the calculator, get the reports. "
                  "Nothing else to set up.").font = Font(italic=True, color=ABU, size=10)

    baris = [
        ("", ""),
        ("HOW TO USE", ""),
        ("1. Sheet 'D'", "Paste your DEPOSIT rows starting at row 2. Keep the header row as it is."),
        ("2. Sheet 'W'", "Paste your WITHDRAWAL rows starting at row 2. Keep the header row as it is."),
        ("3. Sheet 'Xero'", "Add any missing dates. One row per date, one column per currency, "
                            "quoted as local currency per 1 USD (e.g. VND 26,317.2 = 1 USD)."),
        ("4. Sheet 'Fund Transfer Table'",
         "One row per fund transfer between channels. This feeds the 'Channel Balance' and "
         "'J Wallet (calc)' sheets. Put 'J Wallet' in the recipient column when the money "
         "goes into the USDT wallet."),
        ("5. Sheet 'Opening Balance'",
         "Closing balance of the last day before this month, per currency and channel. "
         "isi_template.py fills it in for you; check it only if a balance looks wrong. "
         "Leave it empty and the running balances simply start from zero."),
        ("6. Run", "Double-click 'Hitung DW.app' and pick this file."),
        ("7. Result", "A new file '<name>-hasil.xlsx' appears next to this one. Your file is never modified."),
        ("", ""),
        ("WHAT YOU GET", ""),
        ("Channel Balance", "Per date, currency and channel: deposits, deposit charges, "
                            "fund transfers, transfer charges, withdrawals and a running "
                            "balance. 'Withdrawal charges' is computed (withdrawal amount x "
                            "withdrawal rate) but stays YELLOW: on a few channels the original "
                            "report shows a larger figure that is still unexplained."),
        ("J Wallet (calc)", "USDT wallet ledger built from Fund Transfer Table, with a "
                            "running balance."),
        ("D&W Report", "Summary by Currency x Payment Gateway - Deposit section and Withdrawal "
                       "section, each with a Grand Total."),
        ("D&W Detail", "One flat table combining D and W, one row per transaction, plus audit columns."),
        ("MTOATD", "Daily MT4 / Wallet / CRM / TD reconciliation, one 51-row block per date, "
                   "plus a month-to-date block on the right. Everything that has a formula is "
                   "computed from D and W - nothing to type in. Rows that have no formula "
                   "(manual breakdowns such as voided orders or rounding differences) are left "
                   "EMPTY and coloured grey on purpose."),
        ("Missing Data", "Everything the calculation could not find: exchange-rate dates missing from "
                         "'Xero', fee rates missing from 'D&W FEE', and empty fields in the source data."),
        ("Legend", "What each calculated column means and what the cell colours mean."),
        ("", ""),
        ("THE PURPLE COLUMNS", ""),
        ("Handling Fee", "Transaction x fee rate from 'D&W FEE', matched on Currency + Payment Gateway. "
                         "Deposit rows use the Deposit rate, withdrawal rows the Withdrawal rate."),
        ("Xero Rate", "Exchange rate from sheet 'Xero' on the transaction date."),
        ("Xero USD", "Transaction / Xero Rate."),
        ("Forex Gain/Loss", "Deposit: Xero USD - USD.   Withdrawal: USD - Xero USD."),
        ("", "Leave these four columns EMPTY - the calculator fills them in."),
        ("", ""),
        ("SHEET 'D&W FEE'", ""),
        ("What it is", "One single fee table, already merged from 'D&W TD FEE' and "
                       "'D&W TD FEE -add.'. This is the only place to maintain fee rates."),
        ("Adding a rate", "Add a row: Currency, Payment Gateway, Deposit, Withdrawal. "
                          "Write rates as percentages (2.70%) or decimals (0.027)."),
        ("Fixed / minimum", "Use 'Deposit Fixed' / 'Withdrawal Fixed' for a flat amount added on top, "
                            "and 'Deposit Min' / 'Withdrawal Min' for a minimum charge. "
                            "Formula: fee = max(amount x percent, min) + fixed."),
        ("Gateway names", "Punctuation and spacing are ignored when matching, so '77 Pay' and '77Pay' "
                          "are the same. Currency suffixes are dropped too: 'PA-PHP' matches 'PA'."),
        ("Missing a rate", "The 'Handling Fee' cell is left EMPTY and coloured RED - never 0 - so a "
                           "missing rate can never quietly understate a total."),
    ]
    for judul, isi in baris:
        r += 1
        if judul and not isi:
            sel = ws.cell(r, 2, judul)
            sel.font = Font(bold=True, size=12, color=BIRU)
        elif judul or isi:
            ws.cell(r, 2, judul).font = Font(bold=True, size=10)
            sel = ws.cell(r, 3, isi)
            sel.font = Font(size=10)
            sel.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 28 if len(isi) > 95 else 15

    # ---------------- D & W ----------------
    for nama, kolom, warna in (("D", HDR_D, BIRU), ("W", HDR_W, HIJAU)):
        ws = out.create_sheet(nama)
        ws.sheet_properties.tabColor = warna
        tulis_header(ws, kolom, warna)
        for c, t in enumerate(kolom, start=1):
            if "DATE" in t.upper():
                ws.cell(2, c).number_format = FMT_TGL
            elif "TIME" in t.upper():
                ws.cell(2, c).number_format = FMT_JAM
            if t in WAJIB:
                ws.cell(1, c).comment = openpyxl.comments.Comment(
                    "Required column - do not rename or delete.", "template", height=70, width=220)

    # ---------------- Xero ----------------
    ws = out.create_sheet("Xero")
    ws.sheet_properties.tabColor = "C55A11"
    if kurs:
        for i, row in enumerate(kurs, start=1):
            for j, v in enumerate(row, start=1):
                sel = ws.cell(i, j, v)
                if i == 1:
                    sel.font = Font(bold=True, color="FFFFFF", size=10)
                    sel.fill = PatternFill("solid", fgColor="C55A11")
                    sel.alignment = Alignment(horizontal="center")
                elif j == 1:
                    sel.number_format = FMT_TGL
                else:
                    sel.number_format = "#,##0.######"
    else:
        ws.cell(1, 1, "Date").font = Font(bold=True)
    ws.column_dimensions["A"].width = 12
    for c in range(2, (len(kurs[0]) if kurs else 2) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.freeze_panes = "B2"

    # ---------------- Fund Transfer Table ----------------
    ws = out.create_sheet("Fund Transfer Table")
    ws.sheet_properties.tabColor = "BF8F00"
    tipis2 = Side(style="thin", color="D9D9D9")
    for r_, peta, warna in ((2, FTT_R2, "7F6000"), (3, FTT_R3, "BF8F00")):
        for c, t in peta.items():
            sel = ws.cell(r_, c, t)
            sel.font = Font(bold=True, color="FFFFFF", size=10)
            sel.fill = PatternFill("solid", fgColor=warna)
            sel.alignment = Alignment(wrap_text=True, horizontal="center")
            if t in FTT_EN:
                sel.comment = openpyxl.comments.Comment(FTT_EN[t], "template",
                                                        height=80, width=250)
            sel.border = Border(left=tipis2, right=tipis2, top=tipis2, bottom=tipis2)
    ws.cell(1, 2, "One row per fund transfer. Left half = money going out of a channel, "
                  "right half = where it lands. Put 'J Wallet' in 收款人 (recipient) when "
                  "the money goes into the USDT wallet.").font = \
        Font(italic=True, size=9, color=ABU)
    for c in range(1, 19):
        ws.column_dimensions[get_column_letter(c)].width = 9 if c == 1 else 15
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "B4"

    # ---------------- Opening Balance ----------------
    # Saldo penutup hari terakhir bulan SEBELUMNYA. Tanpa ini saldo di sheet
    # 'Channel Balance' dan 'J Wallet (calc)' mulai dari nol.
    ws = out.create_sheet(OB_SHEET)
    ws.sheet_properties.tabColor = "7F6000"
    ws.cell(1, 1, "Closing balance of the LAST DAY BEFORE this month's data. "
                  "isi_template.py fills this in automatically from the source workbook; "
                  "you only need to touch it if a balance is wrong or missing. "
                  "Leave it empty and the balances simply start from zero.").font = \
        Font(italic=True, size=9, color=ABU)
    tulis_header(ws, OB_KOLOM, "7F6000")
    for c, t in enumerate(OB_KOLOM, start=1):
        if t in OB_EN:
            ws.cell(2, c).comment = openpyxl.comments.Comment(OB_EN[t], "template",
                                                              height=80, width=260)
    ws.cell(3, 1).number_format = FMT_TGL
    for c, w in enumerate((12, 10, 22, 18), start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A3"

    # ---------------- D&W FEE ----------------
    ws = out.create_sheet("D&W FEE")
    ws.sheet_properties.tabColor = UNGU
    ws["A1"] = "Dupoin Markets"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = " Handling Fee Rate"
    ws["A2"].font = Font(bold=True)
    ws["D2"] = ("Merged from 'D&W TD FEE' + 'D&W TD FEE -add.' - this is now the single "
                "source of truth for fee rates.")
    ws["D2"].font = Font(italic=True, size=9, color=ABU)

    kol_fee = ["Currency", "Payment Gateway", "Deposit", "Withdrawal",
               "Deposit Fixed", "Withdrawal Fixed", "Deposit Min", "Withdrawal Min",
               "Source", "Note"]
    tipis = Side(style="thin", color="D9D9D9")
    for c, t in enumerate(kol_fee, start=1):
        sel = ws.cell(3, c, t)
        sel.font = Font(bold=True, color="FFFFFF", size=10)
        sel.fill = PatternFill("solid", fgColor=UNGU)
        sel.alignment = Alignment(wrap_text=True, horizontal="center")
        sel.border = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)
    ws.row_dimensions[3].height = 30

    r = 3
    for (cur, _k), v in sorted(fee.items(), key=lambda x: (x[0][0], x[1].get("nama") or "")):
        r += 1
        dep, wd = v.get("deposit"), v.get("withdrawal")
        asal = v.get("asal") or []
        catat = list(v.get("ditimpa") or [])
        if v.get("fixed"):
            catat.append(f"old notation '+{v['fixed']:g}' in the source sheet - not applied")
        nilai = [cur, v.get("nama") or "",
                 dep["pct"] if dep else None, wd["pct"] if wd else None,
                 (dep["fixed"] or None) if dep else None, (wd["fixed"] or None) if wd else None,
                 dep["min"] if dep else None, wd["min"] if wd else None,
                 " + ".join(asal[-2:]) if len(asal) > 1 else (asal[0] if asal else ""),
                 "; ".join(catat) or None]
        for c, val in enumerate(nilai, start=1):
            sel = ws.cell(r, c, val)
            sel.border = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)
            if c in (3, 4):
                sel.number_format = "0.00%"
            elif c in (5, 6, 7, 8):
                sel.number_format = "#,##0.##"
            elif c == 10 and val:
                sel.fill = PatternFill("solid", fgColor="FFF2CC")
                sel.alignment = Alignment(wrap_text=False)
    for kol, w in zip("ABCDEFGHIJ", (10, 24, 10, 12, 13, 15, 12, 14, 34, 60)):
        ws.column_dimensions[kol].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:J{r}"

    out.save(dst)
    mb = dst.stat().st_size / 1024
    print(f"\nSimpan : {dst}  ({mb:,.0f} KB)")
    print(f"         sheet: {', '.join(out.sheetnames)}")
    print(f"         'D&W FEE' berisi {r - 3} kombinasi")
    print("\nTinggal tempel data ke sheet 'D' dan 'W', lalu jalankan kalkulatornya.")


if __name__ == "__main__":
    main()
