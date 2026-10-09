"""Blank Deal Segregator uploads, built in memory so Excel cannot rewrite the ids.

Each workbook has the sheet the parser reads first, then a Guide sheet that
marks required and optional columns. Deal, Order, Position and Login are text,
so Excel does not turn them into 1.41E+08.
"""
import io

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from deal_segregator import FX_SHEET, KOLOM_WAJIB, _KOLOM_PETA, _MT4_WAJIB

HEADER_MT5 = [
    "Deal", "ID", "Order", "Position", "Login", "Name", "Group", "Country",
    "Time", "Type", "Entry", "Symbol", "Volume", "Price", "S / L", "T / P",
    "Reason", "Commission", "Fee", "Swap", "Profit", "Dealer", "Currency",
    "Comment",
]
TEKS_MT5 = ("Deal", "ID", "Order", "Position", "Login")

HEADER_MT4 = [
    "Deal", "Login", "Open Time", "Type", "Symbol", "Volume", "Open Price",
    "S/L", "T/P", "Close Time", "Close Price", "Reason", "Gateway Order",
    "Gateway Volume", "Open Price Delta", "Close Price Delta", "Agent",
    "Commission", "Taxes", "Swap", "Profit", "Points", "Comment", "Currency",
]
TEKS_MT4 = ("Deal", "Login", "Gateway Order")

HEADER_AKUN = ["account", "account_type"]
HEADER_FX = ["date", "Currency", "rate"]

NAMA_BERKAS = {
    "mt5": "MT5 Deals History template.xlsx",
    "mt4": "MT4 Raw Report template.xlsx",
    "accounts": "MT4 account types template.xlsx",
    "fx": "Client Equity FX template.xlsx",
}

CATATAN = (
    "Upload the original MT5 or MT4 export (CSV) directly, or fill this template. "
    "Do not open a CSV in Excel and save it again. Deal and Login must stay full "
    "numbers, not 1.41E+08."
)


def _panduan(header, wajib, catatan_kolom=None):
    wajib = set(wajib)
    baris = [["Column", "Required", "Notes"], ["Format", "", CATATAN]]
    for nama in header:
        status = "required" if nama in wajib else "optional"
        baris.append([nama, status, (catatan_kolom or {}).get(nama, "")])
    return baris


def _tulis(judul, header, contoh, panduan, kolom_teks):
    buku = Workbook()
    lembar = buku.active
    lembar.title = judul
    tebal = Font(bold=True)
    lembar.append(header)
    for sel in lembar[1]:
        sel.font = tebal
    teks = {header.index(nama) + 1 for nama in kolom_teks if nama in header}
    for kolom in teks:
        lembar.column_dimensions[get_column_letter(kolom)].number_format = "@"
    for contoh_baris in contoh:
        lembar.append(list(contoh_baris))
        for kolom in teks:
            sel = lembar.cell(lembar.max_row, kolom)
            if sel.value is not None:
                sel.value = str(sel.value)
                sel.number_format = "@"
    lembar.freeze_panes = "A2"
    lembar.auto_filter.ref = lembar.dimensions
    pandu = buku.create_sheet("Guide")
    for baris in panduan:
        pandu.append(baris)
    pandu.column_dimensions["A"].width = 28
    pandu.column_dimensions["B"].width = 14
    pandu.column_dimensions["C"].width = 88
    pandu["A1"].font = tebal
    pandu["B1"].font = tebal
    pandu["C1"].font = tebal
    pandu.freeze_panes = "A2"
    return buku


def buku_mt5():
    header = HEADER_MT5
    contoh = [
        ["900001", "1", "910001", "920001", "100001", "Example Client",
         "real\\Classic", "Thailand", "2026.10.04 01:02:03", "buy", "out",
         "XAUUSD", 0.1, 2650, 0, 0, "0", -1.5, 0, 0, 12.5, "0", "USD",
         "template example"],
        ["900002", "2", "910002", "920002", "100001", "Example Client",
         "real\\Classic", "Thailand", "2026.10.04 01:05:00", "buy", "in",
         "XAUUSD", 0.1, 2651, 0, 0, "0", -0.5, 0, 0, 0, "0", "USD",
         "opening example"],
        ["900003", "3", "910003", "920003", "100001", "Example Client",
         "real\\Classic", "Thailand", "2026.10.04 03:15:00", "sell", "out",
         "EURUSD", 1, 1.08, 0, 0, "0", -1.5, 0, -0.2, -3, "0", "USD",
         "second closing example"],
    ]
    panduan = _panduan(header, KOLOM_WAJIB, {
        "Deal": "Text. Full deal id, never scientific notation.",
        "Login": "Text. Full account number.",
        "Entry": "in or out. Other values are dropped.",
        "Time": "YYYY.MM.DD HH:MM:SS",
    })
    return _tulis("Deals History", header, contoh, panduan, TEKS_MT5)


def buku_mt4():
    """Header-first layout. The trailing Currency column is optional."""
    header = HEADER_MT4
    kosong = [""] * len(header)

    def isi(**nilai):
        baris = list(kosong)
        for nama, nilai_kolom in nilai.items():
            baris[header.index(nama)] = nilai_kolom
        return baris

    contoh = [
        isi(**{
            "Deal": "800001", "Login": "100001",
            "Open Time": "2026.10.04 01:00:00", "Type": "buy", "Symbol": "XAUUSD",
            "Volume": 0.1, "Open Price": 2650, "Close Time": "2026.10.04 02:00:00",
            "Close Price": 2655, "Agent": 0, "Commission": -2, "Taxes": 0,
            "Swap": 0, "Profit": 10, "Comment": "closed example", "Currency": "USD",
        }),
        isi(**{
            "Deal": "800002", "Login": "100001",
            "Open Time": "2026.10.04 02:10:00", "Type": "sell", "Symbol": "EURUSD",
            "Volume": 1, "Open Price": 1.08, "Close Time": "2026.10.04 04:00:00",
            "Close Price": 1.07, "Agent": 0, "Commission": -1, "Taxes": 0,
            "Swap": 0, "Profit": 4, "Comment": "second closed example",
            "Currency": "USD",
        }),
        isi(**{
            "Deal": "800003", "Login": "100001",
            "Open Time": "2026.10.04 05:00:00", "Type": "balance",
            "Profit": -5, "Comment": "deposit example", "Currency": "USD",
        }),
    ]
    panduan = _panduan(header, _MT4_WAJIB, {
        "Deal": "Text. Full deal id.",
        "Login": "Text. Full account number.",
        "Close Time": "Required for a closed buy/sell. Balance and credit rows are dropped.",
        "Currency": "Optional trailing column. USD or USC. Used when the account-type file does not say.",
        "Type": "buy or sell for a trade. balance and credit are dropped.",
    })
    panduan.insert(2, [
        "Layout", "",
        "Header on the first row (comma export). A Manager Raw Report that starts "
        "with a title line is also accepted. This file is for KVB. DPM is MT5 only.",
    ])
    return _tulis("Raw Report", header, contoh, panduan, TEKS_MT4)


def buku_akun():
    contoh = [
        ["100001", "Classic"],
        ["100002", "CentAccount"],
        ["100003", "Plus"],
        ["100004", "Pro"],
    ]
    panduan = _panduan(HEADER_AKUN, _KOLOM_PETA, {
        "account": "Text. The MT4 login.",
        "account_type": "CentAccount is USC. Classic, Plus, Pro, and every other type are USD.",
    })
    panduan.insert(2, [
        "Layout", "",
        "Optional, KVB only. One row per MT4 account. DPM does not use this file.",
    ])
    return _tulis("Accounts", HEADER_AKUN, contoh, panduan, ("account",))


def buku_fx():
    contoh = [
        ["2026.10.04 00:00:00", "USD", None],
        ["2026.10.04 00:00:00", "USC", 100],
        ["2026.10.04 00:00:00", "EUR", 0.92],
    ]
    panduan = _panduan(HEADER_FX, HEADER_FX, {
        "date": "YYYY.MM.DD HH:MM:SS. Must match the deal date.",
        "Currency": "Same labels as the deals file, for example USD, USC, EUR.",
        "rate": "Units of this currency per 1 USD. Leave USD blank.",
    })
    panduan.insert(2, [
        "Sheet", "",
        f"The parser reads the sheet named {FX_SHEET}. date, Currency and rate "
        "are required. Do not rename them.",
    ])
    return _tulis(FX_SHEET, HEADER_FX, contoh, panduan, ("date",))


def bangun(nama: str) -> Workbook:
    if nama == "mt5":
        return buku_mt5()
    if nama == "mt4":
        return buku_mt4()
    if nama == "accounts":
        return buku_akun()
    if nama == "fx":
        return buku_fx()
    raise KeyError(nama)


def bytes_template(nama: str) -> io.BytesIO:
    buku = bangun(nama)
    buf = io.BytesIO()
    buku.save(buf)
    buku.close()
    buf.seek(0)
    return buf
