#!/usr/bin/env python3
"""KVB web pipeline order, arguments, and failure isolation."""
import base64
import io
import os
import tempfile
import time
from pathlib import Path
from types import SimpleNamespace

TMP = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(TMP / "users.db")
os.environ["DW_JOB_DIR"] = str(TMP / "jobs")
os.environ["DW_ADMIN_USER"] = "boss"
os.environ["DW_ADMIN_PASS"] = "boss-password-1"

import app as A  # noqa: E402

A.app.config["TESTING"] = True
client = A.app.test_client()
token = base64.b64encode(b"boss:boss-password-1").decode()
client.environ_base["HTTP_AUTHORIZATION"] = f"Basic {token}"
client.environ_base["HTTP_ACCEPT"] = "text/html"

assert A.TOOLS["dw"][2] == ("isi_template.py", "hitung_dw.py")
assert A.TOOLS["segregate"][2] == ("deal_segregator.py",)
assert A.KVB_TOOLS["dw"][2] == ("siapkan_kvb.py", "isi_template.py", "hitung_dw.py")

real_run = A.subprocess.run
commands = []


def fake_run(command, **_kwargs):
    commands.append(command)
    output = Path(command[command.index("-o") + 1])
    output.write_bytes(b"PK fake xlsx")
    return SimpleNamespace(returncode=0, stdout="ok\n", stderr="")


def wait(location):
    job_id = location.rstrip("/").rsplit("/", 1)[-1]
    for _ in range(200):
        state = A._read_state(job_id)
        if state and state["state"] in ("done", "failed"):
            return job_id, state
        time.sleep(0.01)
    raise AssertionError("job timeout")


try:
    A.subprocess.run = fake_run
    response = client.post("/kvb/tool/dw", data={
        "file": (io.BytesIO(b"fake"), "KVB Plus.xlsx"),
        "bulan": "8", "tahun": "2026",
    }, content_type="multipart/form-data")
    assert response.status_code == 302, (response.status_code, response.data[:500])
    job_id, state = wait(response.headers["Location"])
    assert state["state"] == "done", state
    assert state["company"] == "kvb"
    assert response.headers["Location"].startswith("/kvb/job/")
    assert [Path(c[1]).name for c in commands] == [
        "siapkan_kvb.py", "isi_template.py", "hitung_dw.py"
    ]
    assert "--bulan" not in commands[0]
    assert commands[1][-3:] == ["--bulan", "2026-08", "--tanpa-extra-fees"]
    assert commands[2][-3:] == ["--period", "2026-08", "--tanpa-extra-fees"]
    assert Path(commands[1][2]) == Path(commands[0][commands[0].index("-o") + 1])
    assert Path(commands[2][2]) == Path(commands[1][commands[1].index("-o") + 1])
    assert client.get(f"/kvb/job/{job_id}").status_code == 200

    commands.clear()
    response = client.post("/tool/dw", data={
        "file": (io.BytesIO(b"fake"), "DPM.xlsx"),
        "bulan": "8", "tahun": "2026",
    }, content_type="multipart/form-data")
    _, state = wait(response.headers["Location"])
    assert state["state"] == "done", state
    assert [Path(c[1]).name for c in commands] == ["isi_template.py", "hitung_dw.py"]
    assert "--tanpa-extra-fees" not in commands[1]   # DPM tetap memakai EXTRA_FEES

    commands.clear()
    def fail_translate(command, **_kwargs):
        commands.append(command)
        return SimpleNamespace(returncode=1, stdout="", stderr="STOP: sheet Xero not found\n")
    A.subprocess.run = fail_translate
    response = client.post("/kvb/tool/dw", data={
        "file": (io.BytesIO(b"fake"), "broken.xlsx"),
        "bulan": "8", "tahun": "2026",
    }, content_type="multipart/form-data")
    failed_id, state = wait(response.headers["Location"])
    assert state["state"] == "failed"
    assert len(commands) == 1 and Path(commands[0][1]).name == "siapkan_kvb.py"
    page = client.get(f"/kvb/job/{failed_id}")
    assert page.status_code == 422 and b"sheet Xero not found" in page.data
finally:
    A.subprocess.run = real_run

print("OK KVB web pipeline")
