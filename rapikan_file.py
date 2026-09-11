#!/usr/bin/env python3
"""Perkecil & percepat workbook Excel tanpa mengubah datanya.

Dua sumber pemborosan yang dibersihkan:

1. BARIS HANTU -- Excel menyimpan ratusan ribu baris kosong yang pernah
   diformat/pernah berisi rumus. Contoh nyata: sheet 'D' tercatat 918.498
   baris padahal isinya 182; XML-nya 178 MB.

2. PIVOT CACHE -- kalau file punya PivotTable, Excel ikut menyimpan salinan
   seluruh data sumbernya. Ini bisa 30 MB dan menyumbang 99% waktu buka
   (terukur: 144 detik -> 0,2 detik setelah dibuang). Hanya dibuang kalau
   kamu memberi opsi --buang-pivot, karena PivotTable-nya jadi tidak bisa
   di-refresh lagi (nilai yang terakhir tampil tetap ada sebagai angka biasa).

Bekerja langsung di level XML, jadi selesai dalam hitungan detik dan semua
hal lain (rumus, format, sheet lain) tidak disentuh.

Pakai:
    python3 rapikan_file.py "DPM-D&W.xlsx"
    python3 rapikan_file.py "DPM-D&W.xlsx" --buang-pivot
    python3 rapikan_file.py "DPM-D&W.xlsx" -o "bersih.xlsx" --buang-pivot

File asli TIDAK diubah -- hasilnya file baru "<nama>-rapi.xlsx".
"""

import argparse
import html
import re
import sys
import time
import zipfile
from pathlib import Path

RE_ROW = re.compile(rb"<row[^>]*?(?:/>|>.*?</row>)", re.S)
RE_R_ATTR = re.compile(rb'\br="(\d+)"')
RE_ISI = re.compile(rb"<(?:v|f|is|t)[ >]")
RE_DIM = re.compile(rb'<dimension ref="[^"]*"/>')
RE_SHEETNAME = re.compile(rb'<sheet [^>]*name="([^"]*)"[^>]*r:id="rId(\d+)"')


def kolom_ke_huruf(n):
    s = ""
    while n > 0:
        n, sisa = divmod(n - 1, 26)
        s = chr(65 + sisa) + s
    return s or "A"


def bersihkan_sheet(xml):
    """Buang <row> yang tidak punya nilai/rumus di luar baris data terakhir.
    -> (xml_baru, baris_sebelum, baris_sesudah, jumlah_row_dibuang)"""
    rows = list(RE_ROW.finditer(xml))
    if not rows:
        return xml, 0, 0, 0

    nomor = []
    for m in rows:
        r = RE_R_ATTR.search(m.group(0)[:120])
        nomor.append(int(r.group(1)) if r else 0)
    sebelum = max(nomor) if nomor else 0

    terakhir = 0
    for m, no in zip(rows, nomor):
        if RE_ISI.search(m.group(0)):
            terakhir = max(terakhir, no)

    simpan, dibuang = [], 0
    pos = 0
    keluar = bytearray()
    kol_maks = 0
    for m, no in zip(rows, nomor):
        if no <= terakhir:
            keluar += xml[pos:m.end()]
            for c in re.finditer(rb'<c r="([A-Z]+)\d+"', m.group(0)):
                huruf = c.group(1).decode()
                v = 0
                for ch in huruf:
                    v = v * 26 + (ord(ch) - 64)
                kol_maks = max(kol_maks, v)
            simpan.append(no)
        else:
            keluar += xml[pos:m.start()]
            dibuang += 1
        pos = m.end()
    keluar += xml[pos:]
    xml = bytes(keluar)

    if terakhir and kol_maks:
        ref = f'<dimension ref="A1:{kolom_ke_huruf(kol_maks)}{terakhir}"/>'.encode()
        xml = RE_DIM.sub(ref, xml, count=1)
    return xml, sebelum, terakhir, dibuang


def main():
    ap = argparse.ArgumentParser(description="Perkecil & percepat file Excel")
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--buang-pivot", action="store_true",
                    help="buang PivotTable + cache-nya (paling besar efeknya ke kecepatan)")
    args = ap.parse_args()

    src = Path(args.input).expanduser()
    if not src.exists():
        sys.exit(f"File tidak ditemukan: {src}")
    dst = Path(args.output).expanduser() if args.output else src.with_name(f"{src.stem}-rapi.xlsx")
    if dst.resolve() == src.resolve():
        sys.exit("Output tidak boleh sama dengan input.")

    t0 = time.monotonic()
    mb = src.stat().st_size / 1024 / 1024
    print(f"Baca   : {src.name}  ({mb:,.1f} MB)")

    zin = zipfile.ZipFile(src)
    nama_sheet = {}
    try:
        wbxml = zin.read("xl/workbook.xml")
        rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
        for m in RE_SHEETNAME.finditer(wbxml):
            t = re.search(rf'Id="rId{m.group(2).decode()}"[^>]*Target="([^"]+)"', rels)
            if t:
                nama_sheet["xl/" + t.group(1).lstrip("/")] = html.unescape(m.group(1).decode())
    except KeyError:
        pass

    pivot_byte = sum(i.file_size for i in zin.infolist()
                     if i.filename.startswith(("xl/pivotCache/", "xl/pivotTables/")))
    if pivot_byte and not args.buang_pivot:
        print(f"         file ini punya PivotTable + cache {pivot_byte / 1024 / 1024:,.1f} MB "
              f"(tidak terkompresi).")
        print(f"         Itu biasanya penyebab utama file lambat dibuka.")
        print(f"         Tambahkan --buang-pivot kalau PivotTable-nya tidak dipakai lagi.")

    hapus = ("xl/pivotCache/", "xl/pivotTables/") if args.buang_pivot else ()
    total_row = 0
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            nm = it.filename
            if hapus and nm.startswith(hapus):
                continue
            data = zin.read(nm)

            if nm.startswith("xl/worksheets/sheet") and nm.endswith(".xml"):
                data, sebelum, sesudah, dibuang = bersihkan_sheet(data)
                total_row += dibuang
                label = nama_sheet.get(nm, nm.rsplit("/", 1)[-1])
                if dibuang:
                    print(f"  {label!r:22} {sebelum:>9,} baris -> {sesudah:>7,}   "
                          f"({dibuang:,} baris hantu dibuang)")
                else:
                    print(f"  {label!r:22} {sebelum:>9,} baris -> {sesudah:>7,}   (sudah bersih)")

            elif hapus and nm == "[Content_Types].xml":
                data = re.sub(rb'<Override[^>]*PartName="/xl/pivot[^>]*/>', b"", data)
            elif hapus and nm.endswith(".rels"):
                data = re.sub(rb'<Relationship[^>]*Target="[^"]*pivot[^"]*"[^>]*/>', b"", data)
            elif hapus and nm == "xl/workbook.xml":
                data = re.sub(rb"<pivotCaches>.*?</pivotCaches>", b"", data, flags=re.S)

            zout.writestr(it, data)

    mb2 = dst.stat().st_size / 1024 / 1024
    print(f"\nSimpan : {dst.name}  ({mb2:,.1f} MB)")
    if mb:
        print(f"         {mb:,.1f} MB -> {mb2:,.1f} MB  ({(1 - mb2 / mb) * 100:.0f}% lebih kecil)")
    if total_row:
        print(f"         {total_row:,} baris hantu dibuang")
    if args.buang_pivot and pivot_byte:
        print(f"         PivotTable + cache dibuang ({pivot_byte / 1024 / 1024:,.1f} MB)")
    print(f"Selesai dalam {time.monotonic() - t0:.1f} detik")
    print("\nBuka dulu file hasilnya dan periksa datanya. Kalau sudah benar,")
    print("pakai file ini untuk seterusnya -- proses hitung akan jauh lebih cepat.")


if __name__ == "__main__":
    main()
