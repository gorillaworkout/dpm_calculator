"""Smoke test: python3 test_app.py  (butuh 'Template D&W.xlsx' + 'D&W JUN 2026.xlsx')

Model app-nya ASINKRON sejak 28 Aug 2026: POST -> 302 ke /job/<id>, kerjanya di
thread, lalu /job/<id>/download. Jadi tes ini POST, ambil job id dari header
Location, tunggu state 'done'/'failed', baru unduh.
"""
import io
import shutil
import time
from pathlib import Path

from openpyxl import load_workbook

import app as A
from app import JOBS, TOOLS, app

assert app.config["MAX_CONTENT_LENGTH"] == 200 * 1024 * 1024
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024
c = app.test_client()
HERE = Path(__file__).parent
DEALS_FIXTURE = Path(
    "/Users/bayudarmawan/Documents/Dupoin/zern/"
    "31 Jul_Deals History 2026_08_07 12_07_27 (1).csv"
)

# Setiap upload WAJIB menyertakan bulan laporan (28 Aug 2026 -> 31 Aug 2026).
BULAN = {"bulan": "6", "tahun": "2026"}


def kirim(nama_file, fh, **tambahan):
    """POST lalu tunggu sampai job selesai. -> (job_id, state, info)"""
    data = {"file": (fh, nama_file)}
    data.update(tambahan)
    r = c.post("/tool/dw", data=data)
    if r.status_code != 302:
        return None, r.status_code, r
    job_id = r.headers["Location"].rstrip("/").rsplit("/", 1)[-1]
    for _ in range(600):                       # maksimal 5 menit
        info = A._read_state(job_id)
        if info and info["state"] in ("done", "failed"):
            return job_id, info["state"], info
        time.sleep(0.5)
    raise AssertionError("job tidak pernah selesai")


landing = c.get("/")
assert landing.status_code == 200
assert set(TOOLS) == {"dw", "segregate"}
assert b"Dupoin DPM Tools" in landing.data
assert b"Generate D&amp;W" in landing.data
assert b"Deal Segregator" in landing.data
assert c.get("/tool/mtoatd").status_code == 404, "menu yang tidak ada harus 404"
assert c.get("/tool/hitung").status_code == 404, "menu lama sudah dihapus"
assert c.get("/tool/../../etc/passwd").status_code == 404

# Halaman upload harus memuat pemilih bulan, dan default-nya BULAN LALU.
halaman = c.get("/tool/dw")
assert halaman.status_code == 200
assert b'name="bulan"' in halaman.data and b'name="tahun"' in halaman.data, \
    "dropdown bulan laporan hilang dari halaman upload"
assert b'name="saldo_jw"' in halaman.data
assert b"multiple" not in halaman.data

halaman_segregate = c.get("/tool/segregate")
assert halaman_segregate.status_code == 200
assert b'name="file"' in halaman_segregate.data
assert b"multiple" in halaman_segregate.data
assert b'name="bulan"' not in halaman_segregate.data
assert b'name="saldo_jw"' not in halaman_segregate.data
assert b"Download the template" not in halaman_segregate.data

assert c.post("/tool/segregate", data={},
              content_type="multipart/form-data").status_code == 400
assert c.post(
    "/tool/segregate",
    data={"file": [(io.BytesIO(b"x"), "a.txt")]},
    content_type="multipart/form-data",
).status_code == 400


def kirim_segregate(files):
    r = c.post("/tool/segregate", data={"file": files},
               content_type="multipart/form-data")
    assert r.status_code == 302, (r.status_code, r.data[:500])
    job_id = r.headers["Location"].rstrip("/").rsplit("/", 1)[-1]
    for _ in range(600):
        info = A._read_state(job_id)
        if info and info["state"] in ("done", "failed"):
            return job_id, info
        time.sleep(0.5)
    raise AssertionError("Deal Segregator job did not finish")


# Unicode-only stems are stripped by secure_filename; the validated suffix must survive.
captured = {}
unicode_job_id = None
real_thread = A.threading.Thread
class CapturingThread:
    def __init__(self, target, args, **kwargs):
        captured["sources"] = args[3]
    def start(self):
        pass
A.threading.Thread = CapturingThread
try:
    r = c.post("/tool/segregate", data={"file": (io.BytesIO(b"x"), "交易.xlsx")},
               content_type="multipart/form-data")
    assert r.status_code == 302
    unicode_job_id = r.headers["Location"].rstrip("/").rsplit("/", 1)[-1]
    assert captured["sources"][0].suffix == ".xlsx", captured["sources"][0]
finally:
    A.threading.Thread = real_thread
    if unicode_job_id:
        shutil.rmtree(JOBS / unicode_job_id, ignore_errors=True)
        A._state_path(unicode_job_id).unlink(missing_ok=True)

assert DEALS_FIXTURE.is_file(), f"real Deals History fixture missing: {DEALS_FIXTURE}"
job_id, info = kirim_segregate([(DEALS_FIXTURE.open("rb"), DEALS_FIXTURE.name)])
assert info["state"] == "done", info
assert info["name"].endswith("-hasil.xlsx"), info["name"]
r = c.get(f"/job/{job_id}/download")
assert r.status_code == 200
assert r.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
wb = load_workbook(io.BytesIO(r.data), read_only=True, data_only=True)
try:
    assert wb.sheetnames == ["Daily", "Monthly Summary", "Verifikasi"]
    assert wb["Daily"].max_row - 1 == 9418
    assert wb["Monthly Summary"].max_row - 1 == 9418
finally:
    wb.close()

job_id, info = kirim_segregate([
    (DEALS_FIXTURE.open("rb"), DEALS_FIXTURE.name),
    (DEALS_FIXTURE.open("rb"), DEALS_FIXTURE.name),
])
assert info["state"] == "done", info
assert "dipakai: 209838" in (info["log"] or ""), info["log"]
c.get(f"/job/{job_id}/download")

# --- yang harus DITOLAK sebelum job dibuat ---------------------------------
assert c.post("/tool/dw", data={"file": (io.BytesIO(b"x"), "a.txt"), **BULAN}
              ).status_code == 400, "file bukan .xlsx harus ditolak"

src = HERE / "D&W JUN 2026.xlsx"
assert src.is_file(), "butuh 'D&W JUN 2026.xlsx' berisi data untuk tes ini"

r = c.post("/tool/dw", data={"file": (src.open("rb"), src.name)})
assert r.status_code == 400 and b"report month" in r.data, \
    "upload tanpa bulan laporan harus ditolak"
r = c.post("/tool/dw", data={"file": (src.open("rb"), src.name),
                             "bulan": "13", "tahun": "2026"})
assert r.status_code == 400, "bulan 13 harus ditolak"

# --- template KOSONG -> pesan jelas, bukan 500 mentah ----------------------
tpl = HERE / "Template D&W.xlsx"
_id, state, info = kirim(tpl.name, tpl.open("rb"), **BULAN)
assert state == "failed", state
assert "TIDAK ADA BARIS DATA" in (info["log"] or ""), (info["log"] or "")[:400]

# --- file yang ADA datanya -> jadi, dan bulannya masuk nama file -----------
job_id, state, info = kirim(src.name, src.open("rb"), **BULAN)
assert state == "done", (state, info.get("error"), (info.get("log") or "")[-600:])
assert info["periode"] == "2026-06", info["periode"]
assert "Jun 2026" in info["name"], info["name"]
assert "--period 2026-06" in info["log"] or "2026-06" in info["log"]

r = c.get(f"/job/{job_id}/download")
assert r.status_code == 200, r.status_code
assert r.data[:2] == b"PK", r.data[:200]                       # xlsx, bukan HTML
assert "attachment" in r.headers["Content-Disposition"]
assert not (JOBS / job_id).exists(), "job dir harus dibersihkan setelah diunduh"
hasil_jun = len(r.data)

# --- saldo pembuka J Wallet yang diketik di halaman upload -------------------
halaman = c.get("/tool/dw")
assert b'name="saldo_jw"' in halaman.data, "field saldo pembuka hilang dari halaman"
assert b'id="prevlbl"' in halaman.data, "label bulan sebelumnya hilang"

r = c.post("/tool/dw", data={"file": (src.open("rb"), src.name),
                             "saldo_jw": "bukan angka!!", **BULAN})
assert r.status_code == 400 and b"must be a number" in r.data, "saldo ngawur harus ditolak"

job_id, state, info = kirim(src.name, src.open("rb"), saldo_jw="377.233,36", **BULAN)
assert state == "done", (state, info.get("error"), (info.get("log") or "")[-500:])
assert info["saldo_jw"] == "377.233,36", info["saldo_jw"]
# angka Eropa/Indonesia harus dibaca 377233.36, dan tanggalnya akhir bulan SEBELUMNYA
assert "377,233.36" in info["log"], (info["log"] or "")[:1500]
assert "2026-05-31" in info["log"], (info["log"] or "")[:1500]
c.get(f"/job/{job_id}/download")

# tanpa saldo -> tidak boleh ada baris 'Saldo :' sama sekali
job_id, state, info = kirim(src.name, src.open("rb"), **BULAN)
assert state == "done", state
assert info["saldo_jw"] is None
assert "diketik di halaman upload" not in (info["log"] or "")
c.get(f"/job/{job_id}/download")

# --- bulan yang SALAH -> berhenti dengan pesan yang menyebut bulannya --------
# Sengaja GAGAL, bukan menghasilkan workbook kosong: file kosong yang kelihatan
# normal jauh lebih berbahaya daripada pesan error.
job_id, state, info = kirim(src.name, src.open("rb"), bulan="1", tahun="2026")
assert state == "failed", state
assert "BULAN LAPORANNYA SALAH" in (info["error"] or ""), info["error"]
assert "2026-06" in (info["log"] or ""), (info["log"] or "")[-800:]

# semua menu yang terdaftar harus bisa dibuka
for slug in TOOLS:
    assert c.get(f"/tool/{slug}").status_code == 200, slug

sisa = [p for p in JOBS.iterdir() if p.is_dir()]
assert not sisa, f"job dir bocor: {sisa}"

print("OK", len(TOOLS), "menu,", f"{hasil_jun:,}", "bytes hasil Jun 2026")
