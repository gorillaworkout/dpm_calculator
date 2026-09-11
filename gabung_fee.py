#!/usr/bin/env python3
"""Gabungkan sheet fee 'D&W TD FEE' + 'D&W TD FEE -add.' jadi satu sheet 'D&W FEE'.

Dua tabel itu susunan kolomnya berbeda:
    D&W TD FEE        header di baris 3 : Currency | Payment Gateway | Deposit | Withdrawal
    D&W TD FEE -add.  header di baris 2 : No | Payment Gateway | Currency | Deposit | Withdrawal | (ACTIVE)
                      -- perhatikan Payment Gateway ADA DI DEPAN Currency

Rate juga tidak selalu berupa angka. Bentuk yang ditemukan dan cara dibacanya:
    0.027           -> 2,7%
    1.5%+50         -> 1,5% + biaya tetap 50
    +8              -> 0% + biaya tetap 8
    0.45%,min 90    -> 0,45% dengan minimum 90
    3%+50 / 1%+15 / 0.3%+5 / 3.5%+6

Kalau satu kombinasi Currency + Payment Gateway ada di DUA tabel dengan angka
berbeda, yang dipakai adalah tabel '-add.' (lebih baru, punya penanda ACTIVE),
dan angka lama dicatat di kolom Note. Ganti dengan --utama-menang kalau mau
sebaliknya.

Pakai:
    python3 gabung_fee.py "file.xlsx"
    python3 gabung_fee.py "file.xlsx" -o "D&W FEE gabungan.xlsx" --utama-menang
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:
    sys.exit("openpyxl belum terpasang:  python3 -m pip install --user openpyxl")

# nama sheet dicocokkan tanpa peduli huruf besar/kecil & spasi berlebih
SHEET_UTAMA = ("D&W TD FEE", "D&W FEE", "DW FEE")
SHEET_ADD = ("D&W TD FEE -ADD.", "D&W TD FEE -ADD", "D&W FEE -ADD.", "D&W FEE -ADD")

KOLOM = ["Currency", "Payment Gateway", "Deposit", "Withdrawal",
         "Deposit Fixed", "Withdrawal Fixed", "Deposit Min", "Withdrawal Min",
         "Active", "Source", "Note"]


def norm(s):
    return " ".join(str(s or "").strip().upper().split())


def kunci(gw):
    """Kunci pencocokan gateway: buang semua non-alfanumerik.
    '77 Pay' == '77Pay', '1-2-PAY' == '1-2Pay'."""
    return re.sub(r"[^A-Z0-9]", "", norm(gw))


def cari_sheet(wb, kandidat):
    for s in wb.sheetnames:
        if norm(s) in kandidat:
            return s
    return None


def baca_rate(v):
    """-> (persen, tetap, minimum, teks_asli_kalau_bukan_angka)"""
    if v is None or v == "":
        return None, 0.0, None, None
    if isinstance(v, (int, float)):
        return float(v), 0.0, None, None

    t = str(v).strip()
    asli = t
    pers = tetap = minimum = None

    m = re.search(r"MIN\s*([\d.,]+)", t, re.I)
    if m:
        minimum = float(m.group(1).replace(",", ""))
        t = t[:m.start()] + t[m.end():]

    m = re.search(r"([\d.]+)\s*%", t)
    if m:
        pers = float(m.group(1)) / 100
        t = t[:m.start()] + t[m.end():]

    m = re.search(r"\+\s*([\d.,]+)", t)
    if m:
        tetap = float(m.group(1).replace(",", ""))

    if pers is None and tetap is None and minimum is None:
        try:
            return float(t.replace(",", "")), 0.0, None, None
        except ValueError:
            return None, 0.0, None, asli
    return (pers or 0.0), (tetap or 0.0), minimum, asli


def baca_utama(wb, nama):
    ws = wb[nama]
    out = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i <= 3 or len(row) < 4:
            continue
        cur, gw = row[0], row[1]
        if not cur or not gw:
            continue
        extra = next((str(x).strip() for x in row[4:8]
                      if isinstance(x, str) and x.strip().startswith("+")), None)
        out[(norm(cur), kunci(gw))] = {
            "cur": norm(cur), "gw": str(gw).strip(),
            "dep": row[2], "wd": row[3], "extra": extra}
    return out


def baca_add(wb, nama):
    ws = wb[nama]
    out = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i <= 2 or len(row) < 6:
            continue
        gw, cur = row[2], row[3]
        if not cur or not gw:
            continue
        st = row[6] if len(row) > 6 else None
        out[(norm(cur), kunci(gw))] = {
            "cur": norm(cur), "gw": str(gw).strip(),
            "dep": row[4], "wd": row[5], "aktif": norm(st) == "ACTIVE"}
    return out


def sama(a, b):
    if a is None and b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 1e-12
    return str(a).strip() == str(b).strip()


def main():
    ap = argparse.ArgumentParser(description="Gabungkan dua sheet fee jadi satu 'D&W FEE'")
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--utama-menang", action="store_true",
                    help="kalau konflik, pakai angka dari sheet utama (default: -add.)")
    args = ap.parse_args()

    src = Path(args.input).expanduser()
    if not src.exists():
        sys.exit(f"File tidak ditemukan: {src}")
    dst = Path(args.output).expanduser() if args.output else src.with_name("D&W FEE gabungan.xlsx")

    print(f"Baca   : {src.name}")
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    n_utama, n_add = cari_sheet(wb, SHEET_UTAMA), cari_sheet(wb, SHEET_ADD)
    if not n_utama and not n_add:
        sys.exit(f"Sheet fee tidak ketemu. Sheet yang ada: {wb.sheetnames}")
    print(f"         sheet utama : {n_utama!r}")
    print(f"         sheet -add. : {n_add!r}")

    utama = baca_utama(wb, n_utama) if n_utama else {}
    add = baca_add(wb, n_add) if n_add else {}
    wb.close()

    menang_add = not args.utama_menang
    gabung = {}
    konflik = []

    for k in sorted(set(utama) | set(add)):
        u, a = utama.get(k), add.get(k)
        catatan = []

        if u and a:
            bd = []
            if not sama(u["dep"], a["dep"]):
                bd.append(("Deposit", u["dep"], a["dep"]))
            if not sama(u["wd"], a["wd"]):
                bd.append(("Withdrawal", u["wd"], a["wd"]))
            sumber = "keduanya"
            if bd:
                konflik.append((k, u, a, bd))
                pilih = a if menang_add else u
                lain = u if menang_add else a
                asal = "-add." if menang_add else "TD FEE"
                sumber = f"keduanya (konflik -> pakai {asal})"
                for kol, x, y in bd:
                    lama = (x if menang_add else y)
                    catatan.append(f"{kol} di {'TD FEE' if menang_add else '-add.'}: {lama}")
            else:
                pilih = a if menang_add else u
        elif a:
            pilih, sumber = a, "-add. saja"
        else:
            pilih, sumber = u, "TD FEE saja"

        dep_p, dep_f, dep_m, dep_t = baca_rate(pilih["dep"])
        wd_p, wd_f, wd_m, wd_t = baca_rate(pilih["wd"])
        if dep_t:
            catatan.append(f"Deposit ditulis: {dep_t}")
        if wd_t:
            catatan.append(f"Withdrawal ditulis: {wd_t}")

        ex = (u or {}).get("extra")
        if ex:
            catatan.append(f"tabel utama mencantumkan '{ex}' -- BELUM diterapkan, "
                           f"pindahkan ke kolom Fixed kalau memang berlaku")

        gabung[k] = {
            "cur": pilih["cur"],
            "gw": (a or u)["gw"],
            "dep": dep_p, "wd": wd_p,
            "dep_f": dep_f or None, "wd_f": wd_f or None,
            "dep_m": dep_m, "wd_m": wd_m,
            "aktif": "ACTIVE" if (a or {}).get("aktif") else None,
            "sumber": sumber,
            "note": "; ".join(catatan) or None,
        }

    # ---------- tulis ----------
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "D&W FEE"
    ws["A1"] = "Dupoin Markets"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = " Handling Fee Rate"
    ws["A2"].font = Font(bold=True)
    ws["A2"].comment = None
    for c, t in enumerate(KOLOM, start=1):
        sel = ws.cell(3, c, t)
        sel.font = Font(bold=True, color="FFFFFF")
        sel.fill = PatternFill("solid", fgColor="4472C4")
        sel.alignment = Alignment(wrap_text=True)

    r = 3
    for k in sorted(gabung, key=lambda x: (x[0], gabung[x]["gw"])):
        g = gabung[k]
        r += 1
        for c, v in enumerate([g["cur"], g["gw"], g["dep"], g["wd"], g["dep_f"], g["wd_f"],
                               g["dep_m"], g["wd_m"], g["aktif"], g["sumber"], g["note"]],
                              start=1):
            sel = ws.cell(r, c, v)
            if c in (3, 4):
                sel.number_format = "0.0000%"
            elif c in (5, 6, 7, 8):
                sel.number_format = "#,##0.##"
            if g["note"] and c == 11:
                sel.fill = PatternFill("solid", fgColor="FFF2CC")
    for kol, w in zip("ABCDEFGHIJK", (9, 22, 11, 12, 13, 15, 12, 14, 9, 26, 62)):
        ws.column_dimensions[kol].width = w
    ws.freeze_panes = "A4"

    # sheet konflik
    wk = out.create_sheet("Konflik")
    for c, t in enumerate(["Currency", "Payment Gateway", "Kolom",
                           f"di {n_utama}", f"di {n_add}", "Dipakai", "ACTIVE?"], start=1):
        sel = wk.cell(1, c, t)
        sel.font = Font(bold=True, color="FFFFFF")
        sel.fill = PatternFill("solid", fgColor="C00000")
    rr = 1
    for k, u, a, bd in konflik:
        for kol, x, y in bd:
            rr += 1
            dipakai = y if menang_add else x
            for c, v in enumerate([k[0], (a or u)["gw"], kol, x, y, dipakai,
                                   "ACTIVE" if a.get("aktif") else ""], start=1):
                wk.cell(rr, c, v)
    for kol, w in zip("ABCDEFG", (10, 22, 12, 18, 18, 18, 10)):
        wk.column_dimensions[kol].width = w
    wk.freeze_panes = "A2"

    out.save(dst)

    print(f"\n{len(gabung)} kombinasi digabung:")
    print(f"  dari kedua tabel : {len(set(utama) & set(add))}")
    print(f"  hanya '-add.'    : {len(set(add) - set(utama))}")
    print(f"  hanya utama      : {len(set(utama) - set(add))}")
    print(f"  konflik          : {len(konflik)}  -> yang dipakai: "
          f"{'-add.' if menang_add else 'tabel utama'} (lihat sheet 'Konflik')")
    n_teks = sum(1 for g in gabung.values() if g["dep_f"] or g["wd_f"] or g["dep_m"] or g["wd_m"])
    print(f"  ada biaya tetap / minimum : {n_teks}")
    print(f"\nSimpan : {dst}")


if __name__ == "__main__":
    main()
