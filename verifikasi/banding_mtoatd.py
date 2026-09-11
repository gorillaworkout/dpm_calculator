#!/usr/bin/env python3
"""Bandingkan sheet 'MTOATD' hasil vs \"MTOATD (MAY'26)\" di workbook sumber.

Pakai:  python3 banding_mtoatd.py "file-hasil.xlsx" "workbook-sumber.xlsx"
"""
import os, sys
HASIL = sys.argv[1] if len(sys.argv) > 1 else "../D&W JUN 2026-hasil.xlsx"
SUMBER = sys.argv[2] if len(sys.argv) > 2 else os.path.expanduser(
    "~/Downloads/22.08 Bayu - D&W-Dupoin Markets-JUN 2026v3.xlsx")

import openpyxl, datetime as dt
wa=openpyxl.load_workbook(HASIL,read_only=True,data_only=True)
wb=openpyxl.load_workbook(SUMBER,read_only=True,data_only=True)
A=wa['MTOATD']; B=wb["MTOATD (MAY'26)"]
ga=[list(r) for r in A.iter_rows(min_row=3,max_row=50,min_col=1,max_col=22,values_only=True)]
gb=[list(r) for r in B.iter_rows(min_row=3,max_row=50,min_col=1,max_col=22,values_only=True)]
ok=bad=blank=0; probs=[]
for i in range(1,48):
    for j in range(2,22):
        x,y=ga[i][j],gb[i][j]
        lab=gb[i][1]
        if x is None:
            blank+=1
            if y not in (None,0) and abs(float(y))>0.005:
                probs.append(('KOSONG_TAPI_ADA',i,lab,gb[2-2][j] if False else '',y))
            continue
        y=0.0 if y is None else float(y)
        if abs(float(x)-y)<0.02: ok+=1
        else: bad+=1; probs.append((i,lab,round(float(x),2),round(y,2)))
print('cocok',ok,'beda',bad,'kosong(kami)',blank)
print()
print("PATOKAN 7 Sep 2026: cocok 520, beda 23 -> NORMAL, bukan kerusakan.")
print("22 selisih itu SENGAJA: kita hanya menghitung baris berstatus 'finish',")
print("sementara rumus mereka untuk MT4出金/CRM出金 tidak menyaring status sama")
print("sekali (hanya TD出金 yang menyaring). Keputusan user 28 Aug 2026: kalau")
print("Malaysia ikut menghitung refuse, itu kekeliruan mereka.")
print("Baris yang terpengaruh: MT4出金, CRM出金, CRM-TD出金差异, 本日净入金, 月累计.")
print("Kalau angka 'beda' LEBIH DARI 23, berarti ada yang benar-benar rusak.")
for p in probs[:30]: print(p)
