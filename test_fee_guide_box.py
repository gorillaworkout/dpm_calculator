#!/usr/bin/env python3
"""v1.39: fee-sheet 'How to fill rates' box is tidy and never parsed as a rate.

The box lives in columns I-J (KVB) / L-M (DPM), beyond hitung_dw's legacy
'+N' scan of the 4 cells right of Deposit/Withdrawal. Filling sample rows
next to the box must yield exactly those rows, with fixed=0 (the '+8' example
must NOT become a legacy fixed fee).
"""
import tempfile
from pathlib import Path

from openpyxl import load_workbook

import buat_template_kvb as btk
import hitung_dw as h

ROOT = Path(__file__).resolve().parent
TPL = ROOT / "Template KVB Plus.xlsx"


def test_box_layout():
    wb = load_workbook(TPL)
    hf = wb["Handling Fee"]
    assert [hf.cell(1, c).value for c in range(1, 6)] == btk.FEE_KOLOM
    assert hf["F1"].value is None
    assert hf["G1"].value is None and hf["H1"].value is None
    assert hf["I1"].value == "How to fill rates"
    assert hf["I2"].value == "Write" and hf["J2"].value == "Meaning"
    tulis = [hf.cell(r, 9).value for r in range(3, 3 + len(btk.FEE_CARA_ISI))]
    arti = [hf.cell(r, 10).value for r in range(3, 3 + len(btk.FEE_CARA_ISI))]
    assert tulis == [a for a, _ in btk.FEE_CARA_ISI]
    assert arti == [b for _, b in btk.FEE_CARA_ISI]
    assert "+8" in tulis
    # short tip comment only
    assert hf["C1"].comment and "See the box on the right" in hf["C1"].comment.text
    assert len(hf["C1"].comment.text) < 200
    # Guide sheet still documents the rates and points at the box
    guide = "\n".join(str(c.value) for row in wb["Guide"].iter_rows() for c in row if c.value)
    assert "How to fill rates" in guide
    assert "3.5%; >5000 3.8%" in guide
    assert "cell F1" not in guide
    wb.close()


def test_guide_cells_not_parsed_as_fees():
    """Sample rows beside the box -> exactly those fees; '+8' example ignored."""
    import siapkan_kvb as sk

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        src, mid = tmp / "filled.xlsx", tmp / "siap.xlsx"
        wb = load_workbook(TPL)
        hf = wb["Handling Fee"]
        # Row 4 already holds the '+8' example in column I. Put a real fee on
        # the same row in A-E: if the legacy scan reached column I, fixed
        # would become 8.
        hf["A4"], hf["B4"] = "Pay247-KH USD", "USD"
        hf["C4"], hf["D4"] = "3.5%; >5000 3.8%", 0.02
        hf["E4"] = "live row next to +8 example"
        hf["A5"], hf["B5"] = "ChipPay", "CNY"
        hf["C5"], hf["D5"] = 0.004, 0.007
        # Keep D/W/Xero/Transfer minimal so siapkan_kvb accepts the file.
        wb.save(src)
        wb.close()
        sk.proses(src, mid)

        fee_wb = load_workbook(mid, data_only=True)
        assert "D&W FEE" in fee_wb.sheetnames
        # Guide columns were copied across; '+8' still sits on the Pay247 row.
        dw = fee_wb["D&W FEE"]
        rows = list(dw.iter_rows(values_only=True))
        pay_row = next(r for r in rows if r and r[0] and "PAY247" in str(r[0]).upper())
        assert any(isinstance(x, str) and x.strip() == "+8" for x in pay_row), pay_row

        tabel = h.load_fee_table(fee_wb, pakai_extra_fees=False)
        fee_wb.close()

        assert set(tabel) == {("USD", "PAY247KHUSD"), ("CNY", "CHIPPAY")}, tabel
        pay = tabel[("USD", "PAY247KHUSD")]
        assert abs(pay["fixed"] or 0) < 1e-12, f"+8 example leaked as legacy fixed: {pay}"
        assert pay["deposit"]["tiers"] and pay["deposit"]["tiers"][0]["batas"] == 5000
        assert abs(pay["withdrawal"]["pct"] - 0.02) < 1e-12
        chip = tabel[("CNY", "CHIPPAY")]
        assert abs(chip["deposit"]["pct"] - 0.004) < 1e-12
        assert abs(chip["withdrawal"]["pct"] - 0.007) < 1e-12
        assert abs(chip["fixed"] or 0) < 1e-12
        # No phantom gateway from the Write/Meaning labels.
        assert all(k[1] not in ("WRITE", "MEANING", "HOW TO FILL RATES") for k in tabel)


def test_dpm_generator_has_box():
    """buat_template.py wires the same box at column L on D&W FEE."""
    src = Path("buat_template.py").read_text()
    assert "tulis_fee_cara_isi" in src
    assert "kol_awal=12" in src


test_box_layout()
test_guide_cells_not_parsed_as_fees()
test_dpm_generator_has_box()
print("test_fee_guide_box: ok")
