"""Focused regressions for 7 Sep MTOATD/J Wallet corrections."""
import base64
import datetime as dt
import multiprocessing
import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path

from openpyxl import Workbook

os.environ.setdefault("DW_AUTH_DB", str(Path(tempfile.mkdtemp()) / "test-users.db"))
os.environ.setdefault("DW_ADMIN_USER", "tester")
os.environ.setdefault("DW_ADMIN_PASS", "tester-password-1")
AUTH = "Basic " + base64.b64encode(b"tester:tester-password-1").decode()

import app as web
import hitung_dw as h
from mtoatd_spec import Indeks, hitung


def test_job_page_uses_shared_title_validated_slug_and_tool_specific_copy():
    client = web.app.test_client()
    client.environ_base["HTTP_AUTHORIZATION"] = AUTH
    job_id = uuid.uuid4().hex
    web._write_state(job_id, state="running", name="deals-hasil.xlsx", started=0,
                     error=None, log=None, path=None, slug="segregate")
    try:
        page = client.get(f"/job/{job_id}")
        assert page.status_code == 200
        assert b"Processing \xc2\xb7 Dupoin DPM Tools" in page.data
        assert b"classified" in page.data and b"Client Equity FX" in page.data
        assert b"return to this same job URL" in page.data

        web._write_state(job_id, state="collected", name="deals-hasil.xlsx", started=0,
                         path=None, slug="segregate")
        page = client.get(f"/job/{job_id}")
        assert b'href="/tool/segregate"' in page.data
    finally:
        web._state_path(job_id).unlink(missing_ok=True)


def test_dw_job_processing_copy_remains_accurate():
    client = web.app.test_client()
    client.environ_base["HTTP_AUTHORIZATION"] = AUTH
    job_id = uuid.uuid4().hex
    web._write_state(job_id, state="running", name="dw-hasil.xlsx", started=0,
                     error=None, log=None, path=None, slug="dw")
    try:
        page = client.get(f"/job/{job_id}")
        assert b"workbook is being calculated" in page.data
        assert b"combined, filtered, and deduplicated" not in page.data
    finally:
        web._state_path(job_id).unlink(missing_ok=True)


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


def test_comma_delimited_channel_balance_keeps_thousands_separators():
    balances, count = h.baca_tabel_saldo_channel("77Pay,VND,1,058,257.96")
    assert count == 1
    assert balances[("VND", h.kunci_channel("VND", "77Pay"))] == 1058257.96


def test_combined_channel_and_currency_does_not_inherit_previous_channel():
    balances, count = h.baca_tabel_saldo_channel(
        "Beckpay\tVND\t100\nExamplePay LKR\t250\n\tUSD\t50"
    )
    assert count == 3
    assert balances[("VND", h.kunci_channel("VND", "Beckpay"))] == 100
    assert balances[("LKR", h.kunci_channel("LKR", "ExamplePay"))] == 250
    assert balances[("USD", h.kunci_channel("USD", "ExamplePay"))] == 50
    assert ("EXAMPLEPAY LKR", h.kunci_channel("EXAMPLEPAY LKR", "Beckpay")) not in balances


def test_numeric_yyyymmdd_requires_an_integer_value():
    assert h.as_date(20260703) == dt.date(2026, 7, 3)
    assert h.as_date(20260703.0) == dt.date(2026, 7, 3)
    assert h.as_date(20260703.9) is None


def test_sweeper_never_deletes_an_active_job_even_when_old():
    jobs = Path(tempfile.mkdtemp())
    original_jobs = web.JOBS
    web.JOBS = jobs
    job_id = uuid.uuid4().hex
    job = jobs / job_id
    job.mkdir()
    source = job / "large-source.csv"
    source.write_text("still processing")
    web._write_state(job_id, state="running", started=time.time() - 7200,
                     name="result.zip", slug="segregate", owner_pid=os.getpid())
    old = time.time() - 7200
    os.utime(job, (old, old))
    os.utime(web._state_path(job_id), (old, old))
    try:
        web._sweep_old_jobs(max_age_hours=1)
        assert source.exists()
        assert web._state_path(job_id).exists()
    finally:
        web.JOBS = original_jobs
        shutil.rmtree(jobs, ignore_errors=True)


def test_sweeper_removes_orphaned_running_job():
    jobs = Path(tempfile.mkdtemp())
    original_jobs = web.JOBS
    web.JOBS = jobs
    job_id = uuid.uuid4().hex
    job = jobs / job_id
    job.mkdir()
    (job / "source.csv").write_text("orphaned")
    web._write_state(job_id, state="running", started=time.time() - 7200,
                     name="result.zip", slug="segregate", owner_pid=99999999)
    old = time.time() - 7200
    os.utime(job, (old, old))
    os.utime(web._state_path(job_id), (old, old))
    try:
        web._sweep_old_jobs(max_age_hours=1)
        assert not job.exists()
        assert not web._state_path(job_id).exists()
    finally:
        web.JOBS = original_jobs
        shutil.rmtree(jobs, ignore_errors=True)


def test_completed_job_cleanup_keeps_only_download_result():
    job = Path(tempfile.mkdtemp())
    source = job / "five-gb-source.csv"
    intermediate = job / "step0.xlsx"
    result = job / "result.zip"
    source.write_text("source")
    intermediate.write_text("intermediate")
    result.write_text("result")
    (job / "nested").mkdir()
    (job / "nested" / "temporary").write_text("temporary")
    try:
        web._cleanup_completed_job(job, result)
        assert sorted(p.name for p in job.iterdir()) == ["result.zip"]
        assert result.read_text() == "result"
    finally:
        shutil.rmtree(job, ignore_errors=True)


def test_download_streams_result_without_reading_it_all_into_memory():
    jobs = Path(tempfile.mkdtemp())
    original_jobs = web.JOBS
    original_read_bytes = Path.read_bytes
    web.JOBS = jobs
    job_id = uuid.uuid4().hex
    job = jobs / job_id
    job.mkdir()
    result = job / "result.zip"
    result.write_bytes(b"PK-streamed")
    web._write_state(job_id, state="done", path=str(result), name="result.zip",
                     slug="segregate", started=time.time())

    def forbidden(_path):
        raise AssertionError("download must not call Path.read_bytes()")

    Path.read_bytes = forbidden
    try:
        response = web.app.test_client().get(
            f"/job/{job_id}/download", headers={"Authorization": AUTH})
        assert response.status_code == 200
        assert response.data == b"PK-streamed"
        assert web._read_state(job_id)["state"] == "collected"
        assert not job.exists()
        response.close()
        assert not (jobs / f".{job_id}.download").exists()
    finally:
        Path.read_bytes = original_read_bytes
        web.JOBS = original_jobs
        shutil.rmtree(jobs, ignore_errors=True)


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
        response = web.app.test_client().get(
            f"/job/{job_id}/download", headers={"Authorization": AUTH})
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


if __name__ == "__main__":
    # Tanpa pytest: file ini dijalankan langsung (python3 test_regressions.py),
    # jadi test-nya harus dipanggil sendiri -- kalau tidak, exit 0 itu palsu.
    tests = [v for k, v in sorted(globals().items())
             if k.startswith("test_") and callable(v)]
    for t in tests:
        t()
    print(f"OK {len(tests)} regressions")
