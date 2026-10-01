#!/usr/bin/env python3
"""Session login checks. Run: python3 test_login.py"""
import os
import tempfile
from pathlib import Path

os.environ["DW_AUTH_DB"] = str(Path(tempfile.mkdtemp()) / "t.db")
os.environ["DW_ADMIN_USER"] = "boss"
os.environ["DW_ADMIN_PASS"] = "bosspassword1"

from app import app  # noqa: E402

app.config["TESTING"] = True
HTML = {"Accept": "text/html,application/xhtml+xml"}


def client():
    return app.test_client()


def sign_in(c, user="boss", pw="bosspassword1", **extra):
    return c.post("/login", data={"username": user, "password": pw, **extra})


# --- a browser is sent to the sign-in page, never a Basic Auth popup --------
c = client()
r = c.get("/", headers=HTML)
assert r.status_code == 302 and "/login" in r.headers["Location"], r.headers.get("Location")
assert "WWW-Authenticate" not in r.headers, "the browser would still show the ugly popup"

r = c.get("/login")
assert r.status_code == 200
assert b"Sign in" in r.data and b'name="password"' in r.data
assert b"Dupoin DPM Tools" in r.data

# where the user was heading is remembered
r = c.get("/tool/segregate", headers=HTML)
assert "next=/tool/segregate" in r.headers["Location"], r.headers["Location"]

# --- scripts still get a plain 401, not an HTML redirect -------------------
r = client().get("/tool/segregate")
assert r.status_code == 401 and "WWW-Authenticate" in r.headers
r = client().post("/upload/begin")
assert r.status_code == 401, "the chunked uploader would follow a redirect and break"

# --- wrong credentials ------------------------------------------------------
c = client()
r = sign_in(c, pw="nope")
assert r.status_code == 401 and b"Wrong username or password" in r.data
# the message must not reveal whether the username exists
r2 = sign_in(client(), user="ghost", pw="nope")
assert b"Wrong username or password" in r2.data
assert r.data.replace(b"boss", b"ghost") == r2.data, "the two failures differ"
assert c.get("/", headers=HTML).status_code == 302, "a failed attempt still signed us in"

# --- a good sign-in sets a session and lands on the site --------------------
c = client()
r = sign_in(c)
assert r.status_code == 302 and r.headers["Location"] in ("/", "http://localhost/"), r.headers["Location"]
cookie = r.headers.get("Set-Cookie", "")
assert "dw_session=" in cookie
assert "HttpOnly" in cookie and "Secure" in cookie and "Lax" in cookie, cookie
assert c.get("/", headers=HTML).status_code == 200
assert c.get("/tool/segregate", headers=HTML).status_code == 200

# the page shows who is signed in, and offers a way out
body = c.get("/", headers=HTML).data
assert b"Signed in as" in body and b"boss" in body and b"/logout" in body

# --- ?next= is honoured, but only for our own pages ------------------------
c = client()
r = sign_in(c, next="/tool/dw")
assert r.headers["Location"].endswith("/tool/dw"), r.headers["Location"]

for hostile in ("https://evil.example/steal", "//evil.example", "/login"):
    r = sign_in(client(), next=hostile)
    dest = r.headers["Location"]
    assert "evil.example" not in dest and not dest.endswith("/login"), f"open redirect via {hostile}"

# --- signing out ------------------------------------------------------------
c = client()
sign_in(c)
assert c.get("/", headers=HTML).status_code == 200
r = c.get("/logout")
assert r.status_code == 302 and "/login" in r.headers["Location"]
assert c.get("/", headers=HTML).status_code == 302, "still signed in after signing out"

# --- deleting a user revokes an ACTIVE session immediately -----------------
admin = client()
sign_in(admin)
admin.post("/admin/users", data={"username": "temp", "password": "temp-password-1", "can_dpm": "1"})
victim = client()
sign_in(victim, user="temp", pw="temp-password-1")
assert victim.get("/", headers=HTML).status_code == 200
admin.post("/admin/users/temp/delete")
assert victim.get("/", headers=HTML).status_code == 302, "a deleted user kept working via their cookie"

# --- Basic Auth still works for scripts ------------------------------------
import base64  # noqa: E402
tok = base64.b64encode(b"boss:bosspassword1").decode()
assert client().get("/", headers={"Authorization": f"Basic {tok}"}).status_code == 200

# --- the admin page is reachable and gated ----------------------------------
admin = client()
sign_in(admin)
assert admin.get("/admin", headers=HTML).status_code == 200
admin.post("/admin/users", data={"username": "plain", "password": "plain-password-1", "can_dpm": "1"})
plain = client()
sign_in(plain, user="plain", pw="plain-password-1")
assert plain.get("/admin", headers=HTML).status_code == 403
assert client().get("/admin", headers=HTML).status_code == 302, "admin page open to strangers"

print("OK login")
