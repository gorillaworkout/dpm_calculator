"""Focused regressions for 7 Sep MTOATD/J Wallet corrections."""
import datetime as dt
import multiprocessing
import shutil
import tempfile
import uuid
from pathlib import Path

from openpyxl import Workbook

import app as web
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


def _download_worker(jobs, job_id, start, results):
    web.JOBS = Path(jobs)
    original = Path.read_bytes
    def slow_read(path):
        if path.name == "hasil.xlsx":
            import time
            time.sleep(0.2)
        return original(path)
    Path.read_bytes = slow_read
    start.wait()
    try:
        response = web.app.test_client().get(f"/job/{job_id}/download")
        results.put((response.status_code, response.data[:2]))
    except Exception as exc:
        results.put(("error", type(exc).__name__))


def test_concurrent_download_has_one_winner_and_controlled_loser():
    jobs = Path(tempfile.mkdtemp())
    job_id = uuid.uuid4().hex
    job = jobs / job_id
    job.mkdir()
    result = job / "hasil.xlsx"
    workbook = Workbook()
    workbook.save(result)
    workbook.close()
    original_jobs = web.JOBS
    web.JOBS = jobs
    web._write_state(job_id, state="done", path=str(result), name="hasil.xlsx",
                     slug="segregate", started=0)
    start = multiprocessing.Event()
    results = multiprocessing.Queue()
    workers = [multiprocessing.Process(target=_download_worker,
                args=(jobs, job_id, start, results)) for _ in range(2)]
    try:
        for worker in workers:
            worker.start()
        start.set()
        outcomes = [results.get(timeout=5) for _ in workers]
        for worker in workers:
            worker.join(timeout=5)
        assert sorted(status for status, _ in outcomes) in ([200, 409], [200, 410]), outcomes
        assert next(body for status, body in outcomes if status == 200) == b"PK"
        assert web._read_state(job_id)["state"] == "collected"
        assert not job.exists()
    finally:
        web.JOBS = original_jobs
        for worker in workers:
            if worker.is_alive():
                worker.terminate()
            worker.join()
        shutil.rmtree(jobs, ignore_errors=True)
