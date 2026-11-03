#!/usr/bin/env python3
"""Partner Approval (Beta): access, CSRF, request flow, approval, signing link,
signatures, uploads, SMTP. Runs without LibreOffice (conversion is faked,
except for the 'LibreOffice missing' case, which uses the real function).

python3 test_partner_approval.py
"""
import base64
import io
import os
import re
import sqlite3
import time
from pathlib import Path

import testkit as T

TEST_ROOT = T.hermetic_env("boss", "bosspassword1")

from app import app  # noqa: E402
import partner_approval as PA  # noqa: E402
from partner_approval import docs, mailer  # noqa: E402

T.assert_isolated(TEST_ROOT)
assert PA.SETTINGS.db == TEST_ROOT / "approval.db"
assert PA.SETTINGS.files == TEST_ROOT / "approval_files"
assert PA.SETTINGS.link_days == 14
assert PA.SETTINGS.base_url == "https://dw-calculator.gorillaworkout.id"
assert not PA.SETTINGS.smtp_ready

app.config["TESTING"] = True
HTML = {"Accept": "text/html,application/xhtml+xml"}

REAL_CONVERT = docs.docx_to_pdf
CONVERTED = []


def fake_convert(docx_path, pdf_path, soffice="", timeout=180):
    CONVERTED.append(Path(docx_path).name)
    Path(pdf_path).write_bytes(b"%PDF-1.4\n% fake test pdf\n")
    return Path(pdf_path)


docs.docx_to_pdf = fake_convert


def sign_csrf(token):
    with app.test_request_context():
        return PA._sign_csrf(token)


def spost(client, path, data=None, **kw):
    """POST to a public sign form with its per-link CSRF field, like the page does."""
    return client.post(path, data={**(data or {}), "csrf": sign_csrf(path.split("/")[2])}, **kw)


def sign_in(username, password):
    c = app.test_client()
    r = c.post("/login", data={"username": username, "password": password})
    assert r.status_code == 302, (username, r.status_code)
    return c


admin = sign_in("boss", "bosspassword1")
for name, access in (("sales", {"can_dpm": "1"}), ("sales2", {"can_dpm": "1"}),
                     ("kvbonly", {"can_kvb": "1"})):
    r = admin.post("/admin/users", data={"username": name, "password": f"{name}-password-1", **access})
    assert r.status_code == 302, (name, r.status_code)
sales = sign_in("sales", "sales-password-1")
sales2 = sign_in("sales2", "sales2-password-1")
kvb = sign_in("kvbonly", "kvbonly-password-1")
anon = app.test_client()


def csrf(client, path="/approval/new"):
    page = client.get(path).get_data(as_text=True)
    m = re.search(r'name="csrf_token" value="([0-9a-f]{64})"', page)
    assert m, page[:500]
    return m.group(1)


def rows(sql, *args):
    with PA.STORE._con() as con:          # the store creates its schema on first use
        return con.execute(sql, args).fetchall()


def req(rid):
    return PA.STORE.get(rid)


def latest_token(rid):
    return rows("SELECT * FROM tokens WHERE request_id=? ORDER BY created_at DESC LIMIT 1", rid)[0]


AGREEMENT = {
    "request_type": "agreement", "contract_type": "NEW", "sales_team": "Shawn Team",
    "sales_uid": "100200", "role": "CM", "uid": "999001", "partner_name": "Test Partner",
    "passport": "X0000001", "address": "1 Sample Road\nSample City",
    "team_upline": "ISS L\u00ca TH\u00c0NH T\u00c2M 446157 > RM 543936 LODIS",
    "background": "Dummy background.", "allowance_scheme": "Net Deposit : 12%",
    "remarks": "Rebate will be set by superior", "region": "Vietnam",
    "effective_date": "2026-11-01", "onboard_date": "2026-10-09",
    "submission_date": "2026-10-30", "partner_email": "partner@example.com",
    "desk": "VN Desk", "create_agreement": "Yes",
}
DOWNGRADE = {"request_type": "downgrade", "role": "RM", "uid": "999002",
             "partner_name": "Test Two", "effective_date": "2026-11-01",
             "new_position": "MIB", "downgrade_type": "Downgrade CRM Role", "desk": "ASIA Desk"}


def submit(client, form):
    token = csrf(client)
    r = client.post("/approval/new", data={**form, "csrf_token": token, "confirmed": "1"})
    assert r.status_code == 302, (r.status_code, r.get_data(as_text=True)[:800])
    return int(r.headers["Location"].rstrip("/").rsplit("/", 1)[1])


# ------------------------------------------------------------------ access
r = anon.get("/approval", headers=HTML)
assert r.status_code == 302 and "/login" in r.headers["Location"]
assert anon.get("/approval/new").status_code == 401
# KVB-only user: browser is sent to KVB, scripts get 403
r = kvb.get("/approval", headers=HTML)
assert r.status_code == 302 and r.headers["Location"].endswith("/kvb"), r.headers.get("Location")
r = kvb.get("/approval/new", headers=HTML)
assert r.status_code == 302 and r.headers["Location"].endswith("/kvb")
assert kvb.post("/approval/new", data=AGREEMENT).status_code == 403
assert kvb.get("/approval/admin").status_code in (302, 403)
# nav link: only on DPM pages for DPM users, marked Beta
home = sales.get("/").get_data(as_text=True)
assert 'href="/approval"' in home and "Beta" in home
assert 'href="/approval"' not in kvb.get("/kvb").get_data(as_text=True)
assert 'href="/approval"' not in admin.get("/kvb").get_data(as_text=True)

page = sales.get("/approval/new").get_data(as_text=True)
for label in ("Agreement", "Special MIB", "Dowgrade", "Remove - Exit", "Contract Type",
              "Which Sales Team", "Team Upline ONLY", "111111-222222-333333",
              "Apply for Agreement , Special MIB &amp; Downgrade", "Do u want to create agreement?",
              "New Position", "Type of Downgrade", "Desk Accountable", "Beta"):
    assert label in page, label
assert "Attachment" not in page
assert "\U0001F7E2 Frank Team" in page and "\U0001F7E3 Kevvin Team" in page and "\U0001F7E1 Shawn Team" in page

# ------------------------------------------------------------------- CSRF
assert sales.post("/approval/new", data={**AGREEMENT, "confirmed": "1"}).status_code == 400
assert sales.post("/approval/preview", data=AGREEMENT).status_code == 400
other_token = csrf(sales2)                                        # tokens are per user
assert sales.post("/approval/new", data={**AGREEMENT, "confirmed": "1",
                                         "csrf_token": other_token}).status_code == 400
assert rows("SELECT COUNT(*) n FROM requests")[0]["n"] == 0

# ------------------------------------------------------- preview / confirm
tok = csrf(sales)
r = sales.post("/approval/preview", data={**AGREEMENT, "csrf_token": tok})
assert r.status_code == 200 and r.json["ok"], r.json
assert r.json["summary"].startswith("Partnership Agreement\nEffective Date: 11/01/2026\n"
                                    "Onboard Date: 10/09/2026\nTeam: Shawn\nRegion: Vietnam")
r = sales.post("/approval/preview", data={**AGREEMENT, "passport": "", "csrf_token": tok})
assert r.status_code == 400 and "passport" in r.json["errors"]
r = sales.post("/approval/preview", data={**DOWNGRADE, "csrf_token": tok})
assert r.status_code == 200 and r.json["summary"].startswith("Dowgrade"), r.json
# no-JS: first POST shows the confirmation page, nothing saved yet
r = sales.post("/approval/new", data={**AGREEMENT, "csrf_token": tok})
assert r.status_code == 200 and "Please <span>confirm</span>" in r.get_data(as_text=True)
assert rows("SELECT COUNT(*) n FROM requests")[0]["n"] == 0
# server-side validation errors re-render the form
r = sales.post("/approval/new", data={**AGREEMENT, "uid": "abc", "csrf_token": tok, "confirmed": "1"})
assert r.status_code == 400 and "UID must be a whole number" in r.get_data(as_text=True)

# -------------------------------------------------------------- submit
rid = submit(sales, AGREEMENT)
assert req(rid)["status"] == "pending" and req(rid)["created_by"] == "sales"
mine = sales.get("/approval").get_data(as_text=True)
assert "PA-%05d" % rid in mine and "Pending" in mine
assert sales2.get(f"/approval/{rid}").status_code == 404           # not someone else's
assert "PA-%05d" % rid not in sales2.get("/approval").get_data(as_text=True)

# escaping of user input everywhere
xss_rid = submit(sales, {**AGREEMENT, "partner_name": "<script>alert(1)</script>",
                         "background": "<img src=x onerror=alert(1)>", "role": "BM"})
for client, path in ((sales, f"/approval/{xss_rid}"), (admin, f"/approval/admin/{xss_rid}"),
                     (admin, "/approval/admin"), (sales, "/approval")):
    body = client.get(path).get_data(as_text=True)
    assert "<script>alert(1)" not in body and "<img src=x" not in body, path
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in body, path

# ----------------------------------------------------------- admin only
assert sales.get("/approval/admin").status_code == 403
assert sales.get(f"/approval/admin/{rid}").status_code == 403
assert sales.post(f"/approval/admin/{rid}/approve", data={"csrf_token": tok}).status_code == 403
assert req(rid)["status"] == "pending"
atok = csrf(admin)
assert admin.post(f"/approval/admin/{rid}/approve").status_code == 400   # CSRF for admin too
lst = admin.get("/approval/admin").get_data(as_text=True)
assert "Pending" in lst and "Approved (Awaiting Signature)" in lst and "Re-sign requested" in lst
assert "PA-%05d" % rid in admin.get("/approval/admin?status=pending").get_data(as_text=True)
assert "PA-%05d" % rid not in admin.get("/approval/admin?status=completed").get_data(as_text=True)

# --------------------------------------------- approve, SMTP NOT configured
r = admin.post(f"/approval/admin/{rid}/approve", data={"csrf_token": atok})
assert r.status_code == 302
a = req(rid)
assert a["status"] == "approved", a
assert a["email_note"] == "Email not sent: SMTP not configured"
assert CONVERTED == ["agreement.docx"]
t1 = latest_token(rid)
assert len(t1["token"]) >= 43                                  # token_urlsafe(32)
assert abs((t1["expires_at"] - t1["created_at"]) - 14 * 86400) < 5
detail = admin.get(f"/approval/admin/{rid}").get_data(as_text=True)
assert f"https://dw-calculator.gorillaworkout.id/sign/{t1['token']}" in detail
assert "Email not sent: SMTP not configured" in detail and ">Copy<" in detail
# a second approve does nothing
admin.post(f"/approval/admin/{rid}/approve", data={"csrf_token": atok})
assert len(rows("SELECT * FROM tokens WHERE request_id=?", rid)) == 1
audit = admin.get(f"/approval/admin/{rid}").get_data(as_text=True)
assert "WIB" in audit and "boss" in audit and "Submitted" in audit

# -------------------------------------------------- role with no template
r = admin.post(f"/approval/admin/{xss_rid}/approve", data={"csrf_token": atok})
b = req(xss_rid)
assert b["status"] == "approved" and b["agreement_note"] == "Agreement template for BM not ready yet"
assert not rows("SELECT * FROM tokens WHERE request_id=?", xss_rid)
assert "Agreement template for BM not ready yet" in admin.get(f"/approval/admin/{xss_rid}").get_data(as_text=True)

# ---------------------------------------------------------- signing page
anon = app.test_client()
url = f"/sign/{t1['token']}"
r = anon.get(url)
assert r.status_code == 200
assert r.headers["Cache-Control"] == "no-store" and r.headers["Referrer-Policy"] == "no-referrer"
body = r.get_data(as_text=True)
for s in ("Test Partner", "X0000001", "Review", "Draw", "Type name", "Upload signed copy",
          "Undo", "Clear", "I have read and agree", "Open PDF", "Download PDF", 'name="csrf"'):
    assert s in body, s
r = anon.get(url + "/agreement.pdf")
assert r.status_code == 200 and r.data.startswith(b"%PDF")
assert anon.get("/sign/not-a-real-token").status_code == 404
assert anon.get("/sign/" + "x" * 300).status_code == 404

from PIL import Image, ImageDraw  # noqa: E402


def drawn_data_url():
    img = Image.new("RGBA", (500, 160), (0, 0, 0, 0))
    ImageDraw.Draw(img).line([(20, 120), (150, 30), (300, 130), (480, 40)], fill=(10, 20, 90, 255), width=4)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


# no "I agree" / blank signature are refused and do not use the link
r = spost(anon, url + "/sign", data={"mode": "draw", "signature": drawn_data_url()})
assert r.status_code == 400 and "I agree" in r.get_data(as_text=True)
blank = Image.new("RGBA", (200, 80), (0, 0, 0, 0))
buf = io.BytesIO()
blank.save(buf, "PNG")
r = spost(anon, url + "/sign", data={"mode": "draw", "agree": "1",
              "signature": "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()})
assert r.status_code == 400 and "draw your signature" in r.get_data(as_text=True)
assert latest_token(rid)["used_at"] is None

# drawn signature
r = spost(anon, url + "/sign", data={"mode": "draw", "agree": "1", "signature": drawn_data_url()},
              headers={"X-Forwarded-For": "203.0.113.7, 10.0.0.1", "User-Agent": "TestBrowser/1.0"})
assert r.status_code == 200 and "Thank you" in r.get_data(as_text=True)
assert "signed.pdf" not in r.get_data(as_text=True)                  # partner gets no file
s = req(rid)
assert s["status"] == "signed" and s["signed_kind"] == "drawn", s
assert s["signer_ip"] == "203.0.113.7" and s["signer_ua"] == "TestBrowser/1.0" and s["signed_at"]
assert (PA.SETTINGS.files / s["signed_file"]).read_bytes().startswith(b"%PDF")
assert (PA.SETTINGS.files / s["signed_docx"]).is_file()
assert docs.unfilled_placeholders(PA.SETTINGS.files / s["signed_docx"], signed=True) == []
assert docs.company_column_blank(PA.SETTINGS.files / s["signed_docx"])
# single use
for path in (url, url + "/agreement.pdf"):
    r = anon.get(path)
    assert "Already submitted" in r.get_data(as_text=True), path
r = spost(anon, url + "/sign", data={"mode": "typed", "agree": "1", "typed_name": "Again"})
assert "Already submitted" in r.get_data(as_text=True) and req(rid)["signed_kind"] == "drawn"

# admin review
d = admin.get(f"/approval/admin/{rid}").get_data(as_text=True)
assert "Signed \u2013 Under Review" in d and "203.0.113.7" in d and "Mark Completed" in d
r = admin.get(f"/approval/admin/{rid}/file/signed_file")
assert r.status_code == 200 and r.data.startswith(b"%PDF") and "attachment" in r.headers["Content-Disposition"]
assert sales.get(f"/approval/admin/{rid}/file/signed_file").status_code == 403
assert admin.get(f"/approval/admin/{rid}/file/../../users.db").status_code == 404
assert admin.get(f"/approval/admin/{rid}/file/data").status_code == 404

# ------------------------------------------- request re-sign, SMTP mocked
SENT = []


class FakeSMTP:
    def __init__(self, host, port, timeout=None):
        self.host, self.port, self.calls = host, port, []

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def starttls(self, context=None):
        self.calls.append("starttls")

    def login(self, user, password):
        self.calls.append(("login", user, password))

    def send_message(self, msg):
        SENT.append((self, msg))


os.environ.update(DW_SMTP_HOST="smtp.test.invalid", DW_SMTP_PORT="2525", DW_SMTP_USER="mailer",
                  DW_SMTP_PASSWORD="pw", DW_SMTP_FROM="Dupoin <no-reply@example.com>",
                  DW_SIGN_LINK_DAYS="3")
PA.reload_settings()
real_smtp = mailer.smtplib.SMTP
mailer.smtplib.SMTP = FakeSMTP
try:
    r = admin.post(f"/approval/admin/{rid}/resign",
                   data={"csrf_token": atok, "note": "Please sign <again> & initial page 2"})
    assert r.status_code == 302
    s = req(rid)
    assert s["status"] == "resign" and s["resign_note"] == "Please sign <again> & initial page 2"
    assert s["email_note"] == "Email sent to partner@example.com"
    t2 = latest_token(rid)
    assert t2["token"] != t1["token"] and abs((t2["expires_at"] - t2["created_at"]) - 3 * 86400) < 5
    (server, msg), = SENT
    assert (server.host, server.port) == ("smtp.test.invalid", 2525)
    assert server.calls == ["starttls", ("login", "mailer", "pw")]
    assert msg["To"] == "Test Partner <partner@example.com>"
    assert msg["Subject"] == "Dupoin Markets \u2013 Partnership Agreement for Signature"
    assert msg["From"] == "Dupoin <no-reply@example.com>"
    text = msg.get_body(("plain",)).get_content()
    html = msg.get_body(("html",)).get_content()
    assert f"/sign/{t2['token']}" in text and f"/sign/{t2['token']}" in html
    assert "Please sign <again> & initial page 2" in text
    assert "Please sign &lt;again&gt; &amp; initial page 2" in html and "<again>" not in html
    # the e-mail failing must not undo anything
    class BrokenSMTP(FakeSMTP):
        def send_message(self, msg):
            raise OSError("connection refused")
    mailer.smtplib.SMTP = BrokenSMTP
    sent, note = mailer.send_signing_link(PA.SETTINGS, "a@example.com", "A", "https://x/sign/t", "x")
    assert not sent and "connection refused" in note
finally:
    mailer.smtplib.SMTP = real_smtp
    for k in ("DW_SMTP_HOST", "DW_SMTP_PORT", "DW_SMTP_USER", "DW_SMTP_PASSWORD",
              "DW_SMTP_FROM", "DW_SIGN_LINK_DAYS"):
        os.environ.pop(k, None)
    PA.reload_settings()
assert not PA.SETTINGS.smtp_ready

url2 = f"/sign/{t2['token']}"
assert "Please sign &lt;again&gt;" in anon.get(url2).get_data(as_text=True)
# the typed tab offers every font in signature_fonts/, each preview in its own @font-face
page = anon.get(url2).get_data(as_text=True)
for fid in ("1-DarceyOliver-Regular", "2-Budaphest-Regular", "3-Augustia"):
    assert f'name="font" value="{fid}"' in page and f"signature_fonts/{fid}.ttf" in page, fid
assert anon.get("/sign/static/signature_fonts/1-DarceyOliver-Regular.ttf").status_code == 200
assert page.count("data-chars=") == 3 and 'id="font-note"' in page
# a typed name no style can show: every style disabled + "draw instead" note (server-rendered)
page = spost(anon, url2 + "/sign", data={"mode": "typed", "agree": "1", "font": "1-DarceyOliver-Regular",
                                       "typed_name": "Nguy\u1ec5n V\u0103n \u0110\u1ee9c"})
assert page.status_code == 400 and "cannot show some letters" in page.get_data(as_text=True)
page = page.get_data(as_text=True)
assert page.count(" disabled>") >= 3 and 'id="font-note" role="status">' in page
assert "Can\u2019t show \u1ec5 \u0110 \u1ee9" in page
assert req(rid)["status"] == "resign" and latest_token(rid)["used_at"] is None
# a name only some styles can show: those are disabled, the first usable one is pre-selected
page = spost(anon, url2 + "/sign", data={"mode": "typed", "font": "2-Budaphest-Regular",
                                       "typed_name": "Jean-Luc O'Brien"}).get_data(as_text=True)
assert 'value="2-Budaphest-Regular" disabled>' in page and 'value="1-DarceyOliver-Regular" checked>' in page
assert 'id="font-note" role="status" hidden>' in page
r = spost(anon, url2 + "/sign", data={"mode": "typed", "agree": "1", "font": "2-Budaphest-Regular",
                                    "typed_name": "Jean-Luc O'Brien"})
assert r.status_code == 400 and "cannot show some letters" in r.get_data(as_text=True)
assert req(rid)["status"] == "resign" and latest_token(rid)["used_at"] is None
# an unknown / path-like font id is refused and nothing is signed
for bad in ("../../app", "3-Augustia.ttf", "/etc/passwd", "nope"):
    r = spost(anon, url2 + "/sign", data={"mode": "typed", "agree": "1", "font": bad,
                                        "typed_name": "L\u00ea Th\u00e0nh T\u00e2m"})
    assert r.status_code == 400 and "Choose one of the signature styles" in r.get_data(as_text=True), bad
    assert req(rid)["status"] == "resign" and latest_token(rid)["used_at"] is None
# typed signature in the chosen style
r = spost(anon, url2 + "/sign", data={"mode": "typed", "agree": "1", "font": "2-Budaphest-Regular",
                                    "typed_name": "Andrew Tan"})
assert r.status_code == 200 and "Thank you" in r.get_data(as_text=True)
s = req(rid)
assert s["status"] == "signed" and s["signed_kind"] == "typed"
assert any("typed name, style Budaphest" in (a["note"] or "") for a in PA.STORE.audit(rid)), PA.STORE.audit(rid)
# under review: compact header actions, textareas only inside the modals, no dark gradients
d = admin.get(f"/approval/admin/{rid}").get_data(as_text=True)
head = d[d.index('class="pa-top"'):d.index('id="next-step"')]
assert "Mark Completed" in head and 'href="#resign-modal"' in head and 'href="#reject-modal"' in head
assert "Signed by the partner" in d and 'id="signed-doc"' in d
body = d[d.index('<main'):]
assert "<textarea" not in body.split('<div class="modal t"')[0]
assert body.count("<textarea") == 2
assert "gradient(" not in d[d.index("/* Partner Approval"):d.index("</style>", d.index("/* Partner Approval"))]
# the signed PDF is shown as page images (admin only, never cached)
from PIL import Image as _Img  # noqa: E402
_pdf = PA.SETTINGS.files / s["signed_file"]
_Img.new("RGB", (595, 842), "white").save(_pdf, "PDF")
img_url = f"/approval/admin/{rid}/signed/page/1.png"
assert img_url in admin.get(f"/approval/admin/{rid}").get_data(as_text=True)
r = admin.get(img_url)
assert r.status_code == 200 and r.data.startswith(docs.PNG_MAGIC) and r.mimetype == "image/png"
assert r.headers["Cache-Control"] == "private, no-store" and r.headers["X-Content-Type-Options"] == "nosniff"
assert admin.get(f"/approval/admin/{rid}/signed/page/2.png").status_code == 404
assert admin.get(f"/approval/admin/{rid}/signed/page/0.png").status_code == 404
assert sales.get(img_url).status_code == 403
assert anon.get(img_url, headers=HTML).status_code in (302, 401)
assert (PA.SETTINGS.files / f"{rid:06d}" / "signature.png").read_bytes().startswith(docs.PNG_MAGIC)
r = admin.post(f"/approval/admin/{rid}/complete", data={"csrf_token": atok})
assert req(rid)["status"] == "completed"
d = admin.get(f"/approval/admin/{rid}").get_data(as_text=True)
assert "Completed" in d and "by boss" in d and "No further actions" in d
assert "Mark Completed" not in d and "#reject-modal" not in d and "<textarea" not in d
assert "Download signed agreement" in d and admin.get(img_url).status_code == 200
r = admin.get(f"/approval/admin/{rid}/file/signed_file")
assert r.status_code == 200

# ------------------------------------------------- expired / revoked links
rid3 = submit(sales, {**AGREEMENT, "partner_name": "Third Partner"})
admin.post(f"/approval/admin/{rid3}/approve", data={"csrf_token": atok})
t3 = latest_token(rid3)["token"]
with sqlite3.connect(PA.SETTINGS.db) as con:
    con.execute("UPDATE tokens SET expires_at=? WHERE token=?", (time.time() - 1, t3))
r = anon.get(f"/sign/{t3}")
assert r.status_code == 410 and "expired" in r.get_data(as_text=True)
r = spost(anon, f"/sign/{t3}/sign", data={"mode": "typed", "agree": "1", "typed_name": "Third Partner"})
assert r.status_code == 410 and req(rid3)["status"] == "approved"
d = admin.get(f"/approval/admin/{rid3}").get_data(as_text=True)
assert "Expired" in d and "The signing link has expired" in d and 'href="#resign-modal"' in d
# admin makes a new link; the old one now says "replaced"
admin.post(f"/approval/admin/{rid3}/resign", data={"csrf_token": atok, "note": ""})
t3b = latest_token(rid3)["token"]
assert t3b != t3
r = anon.get(f"/sign/{t3}")
assert r.status_code == 410 and "replaced" in r.get_data(as_text=True)
assert anon.get(f"/sign/{t3b}").status_code == 200

# ------------------------------------------------------------- uploads
up = f"/sign/{t3b}/upload"
r = spost(anon, up, data={"file": (io.BytesIO(b"MZ\x90\x00 not a pdf"), "signed.pdf")},
              content_type="multipart/form-data")
assert r.status_code == 400 and "PDF, JPG or PNG" in r.get_data(as_text=True)
r = spost(anon, up, data={"file": (io.BytesIO(b"%PDF-" + b"0" * (20 * 1024 * 1024)), "big.pdf")},
              content_type="multipart/form-data")
assert r.status_code == 413
r = spost(anon, up, data={}, content_type="multipart/form-data")
assert r.status_code == 400
assert latest_token(rid3)["used_at"] is None and req(rid3)["status"] == "resign"
# magic bytes alone are not enough: the image must really decode
r = spost(anon, up, data={"file": (io.BytesIO(b"\xff\xd8\xff\xe0 jpeg bytes"), "photo.jpg")},
          content_type="multipart/form-data")
assert r.status_code == 400 and "PDF, JPG or PNG" in r.get_data(as_text=True)
from PIL import Image  # noqa: E402
_jpg = io.BytesIO()
Image.new("RGB", (40, 20), "white").save(_jpg, "JPEG")
r = spost(anon, up, data={"file": (io.BytesIO(_jpg.getvalue()), "photo.exe")},
              content_type="multipart/form-data", headers={"CF-Connecting-IP": "198.51.100.9"})
assert r.status_code == 200 and "Thank you" in r.get_data(as_text=True)
s = req(rid3)
assert s["status"] == "signed" and s["signed_kind"] == "upload" and s["signed_file"].endswith(".jpg")
assert s["signer_ip"] == "198.51.100.9"
assert "Already submitted" in anon.get(f"/sign/{t3b}").get_data(as_text=True)
r = admin.get(f"/approval/admin/{rid3}/file/signed_file?inline=1")
assert r.status_code == 200 and r.data.startswith(b"\xff\xd8\xff")

# --------------------------------------------------- reject needs a reason
rid4 = submit(sales, {**AGREEMENT, "partner_name": "Fourth"})
d = admin.get(f"/approval/admin/{rid4}").get_data(as_text=True)          # pending: compact header
head = d[d.index('class="pa-top"'):d.index('id="next-step"')]
assert 'id="approve-btn"' in head and 'href="#reject-modal"' in head and "Mark Completed" not in head
assert "<textarea" not in d[d.index("<main"):].split('<div class="modal t"')[0]
admin.post(f"/approval/admin/{rid4}/reject", data={"csrf_token": atok, "reason": "  "})
assert req(rid4)["status"] == "pending"
admin.post(f"/approval/admin/{rid4}/reject", data={"csrf_token": atok, "reason": "UID does not match"})
assert req(rid4)["status"] == "rejected" and req(rid4)["reject_reason"] == "UID does not match"
d = admin.get(f"/approval/admin/{rid4}").get_data(as_text=True)
assert 'class="pa-strip bad"' in d and "UID does not match" in d
assert "approve-btn" not in d and "#reject-modal" not in d and "<textarea" not in d
assert "UID does not match" in sales.get(f"/approval/{rid4}").get_data(as_text=True)
assert admin.post(f"/approval/admin/{rid4}/approve", data={"csrf_token": atok}).status_code == 302
assert req(rid4)["status"] == "rejected"

# ------------------------------------------ Downgrade / Exit: no agreement
CONVERTED.clear()
rid5 = submit(sales, DOWNGRADE)
rid6 = submit(sales, {**DOWNGRADE, "request_type": "exit"})
admin.post(f"/approval/admin/{rid5}/approve", data={"csrf_token": atok})
assert req(rid5)["status"] == "completed"
admin.post(f"/approval/admin/{rid6}/reject", data={"csrf_token": atok, "reason": "Not now"})
assert req(rid6)["status"] == "rejected"
assert CONVERTED == [] and not rows("SELECT * FROM tokens WHERE request_id IN (?,?)", rid5, rid6)
assert admin.post(f"/approval/admin/{rid5}/resign", data={"csrf_token": atok}).status_code == 409

# ----------------------------------------- LibreOffice missing: clear error
docs.docx_to_pdf = REAL_CONVERT
PA.SETTINGS.soffice = str(TEST_ROOT / "no-such-soffice")
rid7 = submit(sales, {**AGREEMENT, "partner_name": "Seventh"})
r = admin.post(f"/approval/admin/{rid7}/approve", data={"csrf_token": atok})
s = req(rid7)
assert s["status"] == "approved" and "libreoffice-writer" in s["agreement_note"], s
assert not rows("SELECT * FROM tokens WHERE request_id=?", rid7)
d = admin.get(f"/approval/admin/{rid7}").get_data(as_text=True)
assert "libreoffice-writer" in d and "Retry agreement" in d
docs.docx_to_pdf = fake_convert
PA.SETTINGS.soffice = ""
admin.post(f"/approval/admin/{rid7}/retry", data={"csrf_token": atok})
s = req(rid7)
assert s["agreement_note"] is None and rows("SELECT * FROM tokens WHERE request_id=?", rid7)

# counts per status on the admin list
lst = admin.get("/approval/admin").get_data(as_text=True)
assert re.search(r"Completed<b>2</b>", lst) and re.search(r"Rejected<b>2</b>", lst), lst[-3000:]

T.assert_isolated(TEST_ROOT)
print("test_partner_approval: ok")
