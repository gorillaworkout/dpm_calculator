"""Focused regressions for 7 Sep MTOATD/J Wallet corrections."""
import datetime as dt

import hitung_dw as h
from mtoatd_spec import Indeks, hitung


def test_cross_month_refuse_h1_is_relevant_to_report_period():
    h.PERIODE_FILTER = (2026, 6)
    assert h.refuse_h1_relevan(
        "withdrawal", "refuse", dt.date(2026, 7, 1), dt.date(2026, 6, 30)
    )
    assert not h.refuse_h1_relevan(
        "deposit", "refuse", dt.date(2026, 7, 1), dt.date(2026, 6, 30)
    )


def test_original_currency_adjustment_always_uses_transaction():
    idx = Indeks()
    idx.tambah_withdrawal("USDT", dt.date(2026, 6, 2), dt.date(2026, 6, 1),
                          90, 100, 0, "finish", "Withdrawal")
    assert hitung(idx, "USDT", dt.date(2026, 6, 2))["oc_w_adj_prev"] == -100


def test_completed_date_rows_exclude_refuse():
    idx = Indeks()
    idx.tambah_withdrawal("USDT", dt.date(2026, 6, 18), dt.date(2026, 6, 17),
                          90, 100, 4, "refuse", "Withdrawal", daftar_tanggal=False)
    v = hitung(idx, "USDT", dt.date(2026, 6, 17))
    assert v["crm_w_oc"] == 100
    assert hitung(idx, "USDT", dt.date(2026, 6, 18))["pay_oc"] == 0


def test_empty_notice_does_not_overwrite_opening_balance():
    assert h.baris_pesan_jwallet_kosong(2, 0) == 3
    assert h.baris_pesan_jwallet_kosong(1, 0) == 2
