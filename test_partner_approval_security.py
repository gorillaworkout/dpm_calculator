#!/usr/bin/env python3
"""Partner Approval v1.34 security checks: links, CSRF, rate limits, headers,
state machine, double submit, resend, uploads, e-mail headers, logs.

python3 test_partner_approval_security.py
"""
import io
import logging
import re
import time
from pathlib import Path

import testkit as T

TEST_ROOT = T.hermetic_env("boss", "bosspassword1")

from app import app  # noqa: E402
import partner_approval as PA  # noqa: E402
from partner_approval import docs, forms, mailer, store  # noqa: E402

T.assert_isolated(TEST_ROOT)
app.config["TESTING"] = True


def fake_convert(docx_path, pdf_path, soffice="", timeout=180):
    Path(pdf_path).write_bytes(b"%PDF-1.4\n% fake test pdf\n%%EOF\n")
    return Path(pdf_path)


docs.docx_to_pdf = fake_convert

SENT = []
MAIL_OK = [True]


def fake_send(settings, to, partner_name, link, expires_label, resign_note=None):
    SENT.append((to, link))
    if MAIL_OK[0]:
        return True, f"Email sent to {to}"
    return False, f"Email not sent: SMTP error 535 for {link} password=hunter22secret"


mailer.send_signing_link = fake_send
PA.SETTINGS.smtp["password"] = "hunter22secret"

LOG = io.StringIO()
_h = logging.StreamHandler(LOG)
app.logger.addHandler(_h)
logging.getLogger().addHandler(_h)


def sign_in(username, password):
    c = app.test_client()
    r = c.post("/login", data={"username": username, "password": password})
    assert r.status_code == 302, (username, r.status_code)
    return c


admin = sign_in("boss", "bosspassword1")
for name in ("sales", "sales2"):
    assert admin.post("/admin/users", data={"username": name, "password": f"{name}-password-1",
                                            "can_dpm": "1"}).status_code == 302
sales = sign_in("sales", "sales-password-1")
sales2 = sign_in("sales2", "sales2-password-1")


def csrf(client, path="/approval/new"):
    page = client.get(path).get_data(as_text=True)
    return re.search(r'name="csrf_token" value="([0-9a-f]{64})"', page).group(1)


def sign_csrf(token):
    with app.test_request_context():
        return PA._sign_csrf(token)


AGREEMENT = {
    "request_type": "agreement", "contract_type": "NEW", "sales_team": "Shawn Team",
    "sales_uid": "100200", "role": "CM", "uid": "999001", "partner_name": "Sec Partner",
    "passport": "X0000001", "address": "1 Sample Road", "team_upline": "RM 1",
    "background": "Dummy.", "allowance_scheme": "Net Deposit : 12%", "remarks": "-",
    "region": "Vietnam", "effective_date": "2026-11-01", "onboard_date": "2026-10-09",
    "submission_date": "2026-10-30", "partner_email": "partner@example.com",
    "desk": "VN Desk", "create_agreement": "Yes",
}


def submit(client=None, key=None, **over):
    client = client or sales
    data = {**AGREEMENT, **over, "csrf_token": csrf(client), "confirmed": "1"}
    if key:
        data["submit_key"] = key
    r = client.post("/approval/new", data=data)
    assert r.status_code == 302, (r.status_code, r.get_data(as_text=True)[:500])
    return int(r.headers["Location"].rstrip("/").rsplit("/", 1)[1])


def post(client, path, **data):
    return client.post(path, data={**data, "csrf_token": csrf(client)})


def approve(rid):
    r = post(admin, f"/approval/admin/{rid}/approve")
    assert r.status_code == 302
    tok = PA.STORE.current_token(rid)
    assert tok is not None
    return tok["token"]


def audit(rid):
    return [a["note"] or "" for a in PA.STORE.audit(rid)]


IP = iter(f"203.0.113.{i}" for i in range(1, 250))


def anon_client():
    c = app.test_client()
    ip = next(IP)
    c.environ_base["HTTP_CF_CONNECTING_IP"] = ip
    return c


# ------------------------------------------------------------ token shape
rid = submit()
tok = approve(rid)
assert len(tok) >= 40 and re.fullmatch(r"[A-Za-z0-9_-]+", tok)       # >= 240 bits urlsafe
row = PA.STORE.token(tok)
assert row["expires_at"] - time.time() > 13 * 86400
with PA.STORE._con() as con:
    stored = con.execute("SELECT token_hash FROM tokens WHERE request_id=?", (rid,)).fetchone()
assert stored["token_hash"] == store.token_hash(tok) != tok
assert PA.STORE.token(tok[:-1] + ("A" if tok[-1] != "A" else "B")) is None

# ------------------------------------------------------- sign-page headers
c = anon_client()
r = c.get(f"/sign/{tok}")
assert r.status_code == 200
h = r.headers
assert h["Referrer-Policy"] == "no-referrer" and h["X-Frame-Options"] == "DENY"
assert h["X-Content-Type-Options"] == "nosniff" and h["Cache-Control"] == "no-store"
assert "noindex" in h["X-Robots-Tag"]
csp = h["Content-Security-Policy"]
assert "frame-ancestors 'none'" in csp and "script-src 'nonce-" in csp, csp
nonce = re.search(r"'nonce-([^']+)'", csp).group(1)
body = r.get_data(as_text=True)
assert f'nonce="{nonce}"' in body
assert not re.search(r"\son[a-z]+=", body), "no inline event handlers under the CSP"
assert sign_csrf(tok) in body
assert "Signing page opened" in " ".join(audit(rid))
c.get(f"/sign/{tok}")
assert sum("Signing page opened" in a for a in audit(rid)) == 1      # once per link
r = c.get(f"/sign/{tok}/agreement.pdf")
assert r.status_code == 200 and "frame-ancestors 'self'" in r.headers["Content-Security-Policy"]
assert r.headers["X-Frame-Options"] == "SAMEORIGIN"
# The page shows only this request: no other partner's data, no admin links.
other = submit(partner_name="Other Person Zed", uid="777777")
assert "Other Person Zed" not in body and "/approval/admin" not in body

# ---------------------------------------------------------------- sign CSRF
r = c.post(f"/sign/{tok}/sign", data={"mode": "typed", "typed_name": "Sec Partner", "agree": "1"})
assert r.status_code == 400 and "This page has expired" in r.get_data(as_text=True)
r = c.post(f"/sign/{tok}/sign", data={"mode": "typed", "typed_name": "Sec Partner", "agree": "1",
                                       "csrf": sign_csrf("x" * 43)})
assert r.status_code == 400
r = c.post(f"/sign/{tok}/upload", data={"file": (io.BytesIO(b"%PDF-1.4\n%%EOF\n"), "a.pdf")})
assert r.status_code == 400
assert PA.STORE.get(rid)["status"] == "approved"
# consent is required
r = c.post(f"/sign/{tok}/sign", data={"mode": "typed", "typed_name": "Sec Partner",
                                       "csrf": sign_csrf(tok)})
assert r.status_code == 400 and "tick the box" in r.get_data(as_text=True)
# long and odd input in the typed name is cleaned / capped, never 500
r = c.post(f"/sign/{tok}/sign", data={"mode": "typed", "typed_name": "\u202e" + "A" * 5000,
                                       "agree": "1", "csrf": sign_csrf(tok)})
assert r.status_code in (200, 400), r.status_code

# -------------------------------------------------- uploads: magic + types
rid2 = submit()
tok2 = approve(rid2)
c2 = anon_client()
for name, blob in (("evil.pdf", b"<html><script>alert(1)</script>"),
                   ("x.png", b"\x89PNG\r\n\x1a\n" + b"junk"),
                   ("x.pdf", b"%PDF-1.4 no end marker"),
                   ("x.svg", b"<svg xmlns='http://www.w3.org/2000/svg'/>")):
    r = c2.post(f"/sign/{tok2}/upload", data={"csrf": sign_csrf(tok2),
                                              "file": (io.BytesIO(blob), name)})
    assert r.status_code == 400, (name, r.status_code)
assert docs.check_upload(b"%PDF-1.7\n...\n%%EOF\n") == "pdf"
assert docs.check_upload(b"GIF89a") is None and docs.check_upload(b"") is None
r = c2.post(f"/sign/{tok2}/upload", data={"csrf": sign_csrf(tok2),
                                          "file": (io.BytesIO(b"%PDF-1.4\n1 0 obj\n%%EOF\n"),
                                                   "../../etc/passwd.pdf")})
assert r.status_code == 200 and "Thank you" in r.get_data(as_text=True)
saved = PA.STORE.get(rid2)["signed_file"]
assert saved.endswith("/signed_upload.pdf") and ".." not in saved
assert not str(PA.SETTINGS.files.resolve()).startswith(str(Path(app.static_folder).resolve()))
# single use
r = c2.post(f"/sign/{tok2}/upload", data={"csrf": sign_csrf(tok2),
                                          "file": (io.BytesIO(b"%PDF-1.4\n%%EOF\n"), "a.pdf")})
assert "Already submitted" in r.get_data(as_text=True)
r = c2.get(f"/sign/{tok2}")
assert "Already submitted" in r.get_data(as_text=True)

# ------------------------------------------------ 413 on huge sign bodies
r = c2.post(f"/sign/{tok}/sign", data={"mode": "draw", "agree": "1", "csrf": sign_csrf(tok),
                                        "signature": "data:image/png;base64," + "A" * (4 << 20)})
assert r.status_code == 413 and "too large" in r.get_data(as_text=True)

# ---------------------------------------------------- admin downloads/IDOR
for client in (sales, sales2, app.test_client()):
    r = client.get(f"/approval/admin/{rid2}/file/signed_file")
    assert r.status_code in (302, 401, 403, 404), r.status_code
    assert client.get(f"/approval/admin/{rid2}").status_code in (302, 401, 403, 404)
    if client is not sales2:
        continue
    assert post(client, f"/approval/admin/{other}/approve").status_code in (403, 404)
assert app.test_client().post(f"/approval/admin/{other}/approve").status_code in (302, 401, 403)
assert PA.STORE.get(other)["status"] == "pending"
assert sales2.get(f"/approval/{rid}").status_code in (403, 404)       # not their request
r = admin.get(f"/approval/admin/{rid2}/file/signed_file")
assert r.status_code == 200 and r.headers["X-Content-Type-Options"] == "nosniff"
assert "no-store" in r.headers["Cache-Control"] and r.mimetype == "application/pdf"
assert "attachment" in r.headers["Content-Disposition"]
assert any(a.startswith("Downloaded ") for a in audit(rid2))
assert admin.get(f"/approval/admin/{rid2}/file/../../users.db").status_code == 404
assert admin.get(f"/approval/admin/{rid2}/file/nope").status_code == 404

# ----------------------------------------------------- admin CSRF + state
r = admin.post(f"/approval/admin/{other}/approve")
assert r.status_code in (400, 403)
assert PA.STORE.get(other)["status"] == "pending"
r = post(admin, f"/approval/admin/{other}/complete")                  # pending -> completed
assert PA.STORE.get(other)["status"] == "pending"
r = post(admin, f"/approval/admin/{rid2}/approve")                   # signed -> approve
assert PA.STORE.get(rid2)["status"] == "signed"
r = post(admin, f"/approval/admin/{other}/reject", reason="")
assert PA.STORE.get(other)["status"] == "pending"
for a, b in store.TRANSITIONS.items():
    assert a not in b or a == "resign", (a, b)

# ------------------------------------------------ double approve / race
first = post(admin, f"/approval/admin/{other}/approve")
n_tokens = len([1 for _ in SENT])
again = post(admin, f"/approval/admin/{other}/approve", )
assert again.status_code == 302
page = admin.get(f"/approval/admin/{other}").get_data(as_text=True)
assert "no longer Pending" in page or "already handled" in page
assert len(SENT) == n_tokens                    # no second e-mail
assert not PA.STORE.set_status(other, "approved", "boss", "x", expect=("pending",))

# -------------------------------------------------- submit-key idempotent
key = "k" * 24
a = submit(key=key, partner_name="Dup Check")
b = submit(key=key, partner_name="Dup Check")
assert a == b
assert len([r for r in PA.STORE.list() if r["partner_name"] == "Dup Check"]) == 1
c3 = submit(key="z" * 24, partner_name="Dup Check")
assert c3 != a

# -------------------------------------- SMTP failure: still approved, resend
MAIL_OK[0] = False
rid3 = submit()
tok3 = approve(rid3)
r3 = PA.STORE.get(rid3)
assert r3["status"] == "approved" and r3["email_note"].startswith("Email not sent")
for note in [r3["email_note"], *audit(rid3)]:
    assert "hunter22secret" not in note and tok3 not in note, note
page = admin.get(f"/approval/admin/{rid3}").get_data(as_text=True)
assert "Resend email" in page and "Email not sent" in page
# retry must not create a second link while one is open
before = PA.STORE.current_token(rid3)["token"]
post(admin, f"/approval/admin/{rid3}/retry")
assert PA.STORE.current_token(rid3)["token"] == before
MAIL_OK[0] = True
n = len(SENT)
post(admin, f"/approval/admin/{rid3}/resend")
assert len(SENT) == n + 1 and PA.STORE.current_token(rid3)["token"] == before
assert any(a.startswith("Resend: Email sent") for a in audit(rid3))
post(admin, f"/approval/admin/{rid3}/resend")                          # throttled
assert len(SENT) == n + 1
assert post(sales, f"/approval/admin/{rid3}/resend").status_code in (302, 403, 404)
assert len(SENT) == n + 1

# ---------------------------------------- reject closes the link for good
rid4 = submit()
tok4 = approve(rid4)
post(admin, f"/approval/admin/{rid4}/reject", reason="Not eligible")
c4 = anon_client()
r = c4.get(f"/sign/{tok4}")
assert r.status_code == 410
r = c4.post(f"/sign/{tok4}/sign", data={"mode": "typed", "typed_name": "Sec Partner", "agree": "1",
                                         "csrf": sign_csrf(tok4)})
assert r.status_code == 410 and PA.STORE.get(rid4)["status"] == "rejected"
n = len(SENT)
post(admin, f"/approval/admin/{rid4}/resend")
assert len(SENT) == n

# expired link
rid5 = submit()
tok5 = approve(rid5)
with PA.STORE._con() as con:
    con.execute("UPDATE tokens SET expires_at=1 WHERE request_id=?", (rid5,))
r = anon_client().get(f"/sign/{tok5}")
assert r.status_code == 410 and "expired" in r.get_data(as_text=True)

# ------------------------------------------------------------ XSS
rid6 = submit(partner_name='<img src=x onerror=alert(1)>', remarks="<script>alert(2)</script>",
              address="</textarea><b>x</b>")
tok6 = approve(rid6)
for page in (admin.get(f"/approval/admin/{rid6}").get_data(as_text=True),
             sales.get(f"/approval/{rid6}").get_data(as_text=True),
             anon_client().get(f"/sign/{tok6}").get_data(as_text=True)):
    assert "<img src=x" not in page and "<script>alert" not in page and "<b>x</b>" not in page
subject, text, html = mailer.content(PA.SETTINGS, '<img src=x onerror=alert(1)>',
                                     "https://e.x/sign/abc", "1 Dec 2026",
                                     resign_note="<script>alert(3)</script>")
assert "<img src=x" not in html and "<script>alert" not in html

# ------------------------------------------- e-mail header injection
assert not forms.valid_email("a@b.co\nBcc: x@evil.test")
assert not forms.valid_email("a@b.co\r\n") and not forms.valid_email("a b@c.co")
assert forms.valid_email("first.last+x@example.co.id")
assert mailer.header_safe("Bad\r\nBcc: x@evil.test") == "Bad Bcc: x@evil.test"
msg = mailer.build(PA.SETTINGS, "partner@example.com", "Evil\r\nBcc: x@evil.test",
                   "https://e.x/sign/abc", "1 Dec 2026")
raw = msg.as_string()
assert "\nBcc:" not in raw and msg["Subject"] == mailer.SUBJECT
assert msg.get_content_type() == "multipart/alternative"
r = sales.post("/approval/new", data={**AGREEMENT, "partner_email": "a@b.co\nBcc: x@evil.test",
                                      "csrf_token": csrf(sales), "confirmed": "1"})
assert r.status_code == 400 or "valid" in r.get_data(as_text=True)
assert mailer.scrub("see https://h/sign/" + "Q" * 43 + " pw hunter22secret",
                    "hunter22secret") == "see https://h/sign/*** pw ***"

# ------------------------------------------------------------ rate limits
bad = anon_client()
for i in range(PA.BAD_TOKEN_LIMIT):
    assert bad.get(f"/sign/{'n' * 30}{i}").status_code == 404
r = bad.get(f"/sign/{tok3}")                      # even a good link is blocked now
assert r.status_code == 429 and "Too many attempts" in r.get_data(as_text=True)
assert anon_client().get(f"/sign/{tok3}").status_code == 200   # other IPs unaffected
flood = anon_client()
codes = [flood.post(f"/sign/{tok3}/sign", data={"mode": "typed"}).status_code
         for _ in range(PA.POST_LIMIT + 2)]
assert codes[-1] == 429 and 429 not in codes[:PA.POST_LIMIT]

# ------------------------------------------------------- nothing in logs
logs = LOG.getvalue()
for secret in (tok, tok2, tok3, tok4, tok5, tok6, "hunter22secret"):
    assert secret not in logs, secret

print("test_partner_approval_security: ok")
