#!/usr/bin/env python3
"""Auth + admin page checks. Run: python3 test_auth.py"""
import base64
import os
import tempfile
import threading
from pathlib import Path

os.environ["DW_AUTH_DB"] = str(Path(tempfile.mkdtemp()) / "t.db")
os.environ["DW_ADMIN_USER"] = "boss"
os.environ["DW_ADMIN_PASS"] = "bosspassword1"

from app import app  # noqa: E402  (env must be set before import)
import auth  # noqa: E402

app.config["TESTING"] = True
c = app.test_client()


def hdr(user, pw):
    tok = base64.b64encode(f"{user}:{pw}".encode()).decode()
    return {"Authorization": f"Basic {tok}"}


ADMIN = hdr("boss", "bosspassword1")

# Without credentials every page is refused. Browsers are redirected to the
# sign-in page; anything else gets a plain 401 challenge.
for path in ("/", "/tool/dw", "/tool/segregate", "/template"):
    r = c.get(path)
    assert r.status_code == 401, f"{path} was reachable without a password ({r.status_code})"
    assert "WWW-Authenticate" in r.headers, f"{path} sent no Basic Auth challenge"

r = c.get("/admin")
assert r.status_code == 302 and "/login" in r.headers["Location"], r.status_code

# uploads too -- the expensive endpoint must not be open
r = c.post("/tool/segregate", data={}, content_type="multipart/form-data")
assert r.status_code == 401, "upload endpoint was reachable without a password"

# --- wrong password is rejected ---------------------------------------------
assert c.get("/", headers=hdr("boss", "wrong")).status_code == 401
assert c.get("/", headers=hdr("ghost", "bosspassword1")).status_code == 401

# --- the bootstrap admin works ----------------------------------------------
assert c.get("/", headers=ADMIN).status_code == 200
r = c.get("/admin", headers=ADMIN)
assert r.status_code == 200 and b"User" in r.data

# --- creating a user --------------------------------------------------------
r = c.post("/admin/users", headers=ADMIN,
           data={"username": "siti", "password": "siti-password-1", "can_dpm": "1"})
assert r.status_code == 302, r.data[:400]
assert c.get("/", headers=hdr("siti", "siti-password-1")).status_code == 200

# a normal user cannot reach the admin page or manage anyone
assert c.get("/admin", headers=hdr("siti", "siti-password-1")).status_code == 403
r = c.post("/admin/users", headers=hdr("siti", "siti-password-1"),
           data={"username": "mole", "password": "mole-password-1"})
assert r.status_code == 403, "a normal user was able to create an account"
assert c.get("/", headers=hdr("mole", "mole-password-1")).status_code == 401

# --- validation -------------------------------------------------------------
r = c.post("/admin/users", headers=ADMIN, data={"username": "x", "password": "short"})
assert r.status_code == 400 and b"at least 10" in r.data
r = c.post("/admin/users", headers=ADMIN, data={"username": "siti", "password": "another-one-1", "can_dpm": "1"})
assert r.status_code == 400 and b"already exists" in r.data
r = c.post("/admin/users", headers=ADMIN, data={"username": "bad name!", "password": "valid-password-1"})
assert r.status_code == 400

# --- password reset ---------------------------------------------------------
r = c.post("/admin/users/siti/password", headers=ADMIN, data={"password": "siti-password-2"})
assert r.status_code == 302
assert c.get("/", headers=hdr("siti", "siti-password-1")).status_code == 401, "old password still works"
assert c.get("/", headers=hdr("siti", "siti-password-2")).status_code == 200

# --- deletion ---------------------------------------------------------------
r = c.post("/admin/users/boss/delete", headers=ADMIN)
assert r.status_code == 400 and b"signed in with" in r.data, "admin deleted their own account"
r = c.post("/admin/users/siti/delete", headers=ADMIN)
assert r.status_code == 302
assert c.get("/", headers=hdr("siti", "siti-password-2")).status_code == 401, "deleted user still has access"

# --- passwords are never stored in the clear --------------------------------
raw = Path(os.environ["DW_AUTH_DB"]).read_bytes()
assert b"bosspassword1" not in raw, "the admin password is stored in plain text"

# Multiple Gunicorn workers starting together must all load one winning key.
auth.SECRET_FILE = Path(tempfile.mkdtemp()) / "session.key"
keys = []
threads = [threading.Thread(target=lambda: keys.append(auth.secret_key())) for _ in range(20)]
for thread in threads: thread.start()
for thread in threads: thread.join()
assert len(set(keys)) == 1 and auth.SECRET_FILE.read_text().strip() == keys[0]
assert (auth.SECRET_FILE.stat().st_mode & 0o777) == 0o600

print("OK auth")
