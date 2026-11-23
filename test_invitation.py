#!/usr/bin/env python3
"""DPM Invitation (v1.43): rules, image upload, access, public RSVP form, upsert,
rate limits, honeypot, closed/auto-closed, admin counts, search, export, deletes.

python3 test_invitation.py
"""
import csv
import io
import itertools
import re
import time

import testkit as T

TEST_ROOT = T.hermetic_env("boss", "bosspassword1")

from app import app  # noqa: E402
import invitation as INV  # noqa: E402
from invitation import export, images, rules  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from PIL import Image  # noqa: E402

T.assert_isolated(TEST_ROOT)
assert INV.SETTINGS.db == TEST_ROOT / "invitation.db"
assert INV.SETTINGS.files == TEST_ROOT / "invitation_files"
app.config["TESTING"] = True
HTML = {"Accept": "text/html,application/xhtml+xml"}
_ips = itertools.count(1)


def new_ip():
    return {"X-Real-IP": f"10.0.{next(_ips) // 250}.{next(_ips) % 250}"}


def img_bytes(fmt="JPEG", size=(1024, 1536), exif_gps=False, color=(20, 30, 60)):
    im = Image.new("RGB", size, color)
    buf = io.BytesIO()
    kw = {}
    if exif_gps:
        ex = Image.Exif()
        ex[0x010F] = "SecretCam"                       # Make
        ex[0x8825] = {1: "N", 2: (6.0, 10.0, 0.0)}     # GPS IFD
        kw["exif"] = ex
    im.save(buf, fmt, **kw)
    return buf.getvalue()


# ------------------------------------------------------------------ rules
v, e = rules.parse_response({"full_name": "  Nguyễn   Văn An ", "email": " An@Example.COM ",
                             "attending": "yes", "has_allergy": "no", "allergy_notes": "x"})
assert not e and v == {"full_name": "Nguyễn Văn An", "email": "An@Example.COM",
                       "email_normalized": "an@example.com", "attending": True,
                       "has_allergy": False, "allergy_notes": None}, (v, e)
v, e = rules.parse_response({"full_name": "Bo Li", "email": "bo@x.co", "attending": "no",
                             "has_allergy": "yes", "allergy_notes": "peanuts"})
assert not e and v["attending"] is False and v["has_allergy"] is None and v["allergy_notes"] is None
v, e = rules.parse_response({"full_name": "Bo Li", "email": "bo@x.co", "attending": "yes",
                             "has_allergy": "yes", "allergy_notes": "  shellfish \n gluten "})
assert not e and v["has_allergy"] is True and v["allergy_notes"] == "shellfish gluten"
_, e = rules.parse_response({"full_name": "Bo", "email": "bo@x.co", "attending": "yes", "has_allergy": "yes"})
assert e == {"allergy_notes": "Please tell us what you are allergic to."}, e
_, e = rules.parse_response({"full_name": "Bo", "email": "bo@x.co", "attending": "yes",
                             "has_allergy": "yes", "allergy_notes": "a" * 201})
assert e == {"allergy_notes": "Please use at most 200 characters."}, e
_, e = rules.parse_response({"full_name": "Bo", "email": "bo@x.co", "attending": "yes"})
assert set(e) == {"has_allergy"}, e
_, e = rules.parse_response({"full_name": "B", "email": "a@b", "attending": "maybe"})
assert set(e) == {"full_name", "email", "attending"}, e
_, e = rules.parse_response({"full_name": "x" * 121, "email": "a b@c.d", "attending": "no"})
assert set(e) == {"full_name", "email"}, e
assert rules.slug_issue("dpm-gala-2027") is None
for bad in ("ab", "-abc", "abc-", "a--b", "Abc", "thanks", "a" * 61, "a_b"):
    assert rules.slug_issue(bad), bad
assert rules.suggest_slug("DPM Gala Dinner 2027, Hà Nội!") == "dpm-gala-dinner-2027-ha-noi"
ts, err = rules.parse_closes_at("2030-12-20T18:00", now=0)
assert err is None, err
assert rules.to_wib_input(ts) == "2030-12-20T18:00" and rules.format_wib(ts) == "20 Dec 2030, 18:00 WIB"
import datetime as _dt  # noqa: E402
assert _dt.datetime.fromtimestamp(ts, _dt.timezone.utc).strftime("%Y-%m-%d %H:%M") == "2030-12-20 11:00"
assert rules.parse_closes_at("", now=0) == (None, None)
assert rules.parse_closes_at("2030-02-30T10:00", now=0)[1] == "Enter a valid closing date and time."
assert rules.parse_closes_at("20/12/2030 18:00", now=0)[1] == "Enter a valid closing date and time."
assert "future" in rules.parse_closes_at("2020-01-01T00:00")[1]
assert export._guard("=HYPERLINK(1)") == "'=HYPERLINK(1)" and export._guard("Ann") == "Ann"
print("ok rules")

# ------------------------------------------------------------------ images
for blob, msg in ((b"GIF89a" + b"\0" * 100, "JPG, PNG or WebP"),
                  (img_bytes("GIF"), "JPG, PNG or WebP"),
                  (b"hello, this is text renamed to .jpg", "JPG, PNG or WebP"),
                  (b"\xff" * (images.MAX_BYTES + 1), "larger than 8 MB"),
                  (img_bytes(size=(300, 300)), "too small")):
    try:
        images.process(blob)
    except images.ImageError as exc:
        assert msg in str(exc), (msg, exc)
    else:
        raise AssertionError(msg)
data, info = images.process(img_bytes(exif_gps=True))
assert info["mime"] == "image/webp" and (info["width"], info["height"]) == (1024, 1536), info
assert re.match(images.KEY_RE, info["key"]), info
with Image.open(io.BytesIO(data)) as out:
    assert out.format == "WEBP" and not out.getexif() and "exif" not in out.info
_, info = images.process(img_bytes("PNG", size=(4000, 6000)))
assert (info["width"], info["height"]) == (1600, 2400), info
print("ok images")

# ------------------------------------------------------------------ access
def sign_in(username, password):
    c = app.test_client()
    r = c.post("/login", data={"username": username, "password": password})
    assert r.status_code == 302, (username, r.status_code)
    return c


admin = sign_in("boss", "bosspassword1")
for name, access in (("sales", {"can_dpm": "1"}), ("kvbonly", {"can_kvb": "1"})):
    r = admin.post("/admin/users", data={"username": name, "password": f"{name}-password-1", **access})
    assert r.status_code == 302, (name, r.status_code)
sales = sign_in("sales", "sales-password-1")
kvb = sign_in("kvbonly", "kvbonly-password-1")
anon = app.test_client()


def csrf(client, path="/invitation"):
    html = client.get(path).get_data(as_text=True)
    m = re.search(r'name="csrf_token" value="([0-9a-f]{64})"', html)
    assert m, html[:500]
    return m.group(1)


assert anon.get("/invitation", headers=HTML).status_code == 302
assert anon.get("/invitation").status_code == 401
assert kvb.get("/invitation", headers={"Accept": "*/*"}).status_code == 403
r = sales.get("/invitation")
assert r.status_code == 200 and "No invitations yet" in r.get_data(as_text=True)
assert "New invitation" in r.get_data(as_text=True)                  # every DPM user can create
assert 'id="empty-create"' in r.get_data(as_text=True) and "An admin creates them" not in r.get_data(as_text=True)
assert 'href="/invitation"' in admin.get("/").get_data(as_text=True)  # DPM menu link
assert 'href="/invitation"' not in kvb.get("/kvb").get_data(as_text=True)

ADMIN_TOKEN = re.search(r'name="csrf_token" value="([0-9a-f]{64})"',
                        admin.get("/invitation").get_data(as_text=True)).group(1)
r = sales.post("/invitation/new", data={"csrf_token": "x", "name": "Nope"})
assert r.status_code == 400, r.status_code                            # DPM user: CSRF still checked
r = kvb.post("/invitation/new", data={"csrf_token": "x", "name": "Nope"}, headers={"Accept": "*/*"})
assert r.status_code == 403, r.status_code                            # no Invitation access
r = admin.post("/invitation/new", data={"csrf_token": "bad", "name": "Gala"})
assert r.status_code == 400


def create(name, slug, closes="", blob=None):
    return admin.post("/invitation/new", data={
        "csrf_token": ADMIN_TOKEN, "name": name, "slug": slug, "closes_at": closes,
        "image": (io.BytesIO(blob or img_bytes()), "invite.jpg")}, content_type="multipart/form-data")


r = create("X", "Bad Slug!", "2020-01-01T00:00", img_bytes("GIF"))
assert r.status_code == 422, r.status_code
body = r.get_data(as_text=True)
assert "Enter a name of 2–120 characters." in body and "lowercase letters" in body and "future" in body
assert not any(INV.SETTINGS.files.glob("*")) if INV.SETTINGS.files.exists() else True
r = create("Gala Dinner", "gala", "", img_bytes("GIF"))
assert r.status_code == 422 and "JPG, PNG or WebP" in r.get_data(as_text=True)

r = create("DPM Gala Dinner 2027", "", "2030-12-20T18:00")
assert r.status_code == 303, r.status_code
EID = int(r.headers["Location"].rstrip("/").split("/")[-1])
ev = INV.STORE.event(EID)
assert ev["slug"] == "dpm-gala-dinner-2027" and ev["status"] == "open" and ev["created_by"] == "boss"
assert (INV.SETTINGS.files / ev["image_key"]).is_file()
assert rules.format_wib(ev["closes_at"]) == "20 Dec 2030, 18:00 WIB"
assert create("Another", "dpm-gala-dinner-2027").status_code == 422          # slug taken
SLUG = ev["slug"]
LINK = f"https://dw-calculator.gorillaworkout.id/i/{SLUG}"
for c in (admin, sales):
    page = c.get("/invitation").get_data(as_text=True)
    assert LINK in page and "Copy link" in page and "Open" in page
detail = admin.get(f"/invitation/{EID}")
assert detail.status_code == 200 and detail.headers["Cache-Control"] == "no-store"
assert LINK in detail.get_data(as_text=True) and "Close registration" in detail.get_data(as_text=True)
page = sales.get(f"/invitation/{EID}").get_data(as_text=True)
assert "Close registration" in page and "Save changes" in page          # DPM user manages
assert "Delete this event" not in page                                  # ... but cannot delete it
assert admin.get("/invitation/999").status_code == 404
print("ok access + create")

# ------------------------------------------------------------------ public page
r = anon.get(f"/i/{SLUG}")
html = r.get_data(as_text=True)
assert r.status_code == 200, r.status_code
assert "Confirm your attendance" in html and "Will you attend?" in html and "DPM Gala Dinner 2027" in html
assert "What are you allergic to?" in html and 'name="website"' in html
assert "Please reply before 20 Dec 2030, 18:00 WIB" in html
csp = r.headers["Content-Security-Policy"]
assert "script-src 'nonce-" in csp and "frame-ancestors 'none'" in csp, csp
assert r.headers["Cache-Control"] == "no-store" and r.headers["X-Robots-Tag"] == "noindex, nofollow"
assert anon.get(f"/i/{SLUG.upper()}").status_code == 200
r = anon.get("/i/does-not-exist", headers=HTML)
assert r.status_code == 404 and "This invitation link is not valid" in r.get_data(as_text=True)
r = anon.get(f"/i/{SLUG}/image")
assert r.status_code == 200 and r.mimetype == "image/webp" and "max-age=86400" in r.headers["Cache-Control"]
assert anon.get("/i/does-not-exist/image").status_code == 404
PUB_CSRF = re.search(r'name="csrf" value="([0-9a-f]{64})"', html).group(1)


def submit(data, ip=None, slug=SLUG, token=None):
    return anon.post(f"/i/{slug}", data={"csrf": token or PUB_CSRF, "website": "", **data},
                     headers=ip or new_ip())


r = submit({"full_name": "Ann Lee", "email": "Ann@Dupoin.com", "attending": "yes",
            "has_allergy": "yes", "allergy_notes": "Peanuts"})
assert r.status_code == 303 and r.headers["Location"].endswith(f"/i/{SLUG}/thanks"), r.status_code
r = anon.get(f"/i/{SLUG}/thanks")
assert r.status_code == 200 and "Thank you" in r.get_data(as_text=True) and "Change my answer" in r.get_data(as_text=True)
rows = INV.STORE.responses(EID)
assert len(rows) == 1 and rows[0]["has_allergy"] == 1 and rows[0]["allergy_notes"] == "Peanuts"
# Same email, other case and spaces: updates the same row.
r = submit({"full_name": "Ann Lee-Tan", "email": "  ann@DUPOIN.com ", "attending": "no",
            "has_allergy": "yes", "allergy_notes": "Peanuts"})
assert r.status_code == 303
rows = INV.STORE.responses(EID)
assert len(rows) == 1, rows
assert rows[0]["full_name"] == "Ann Lee-Tan" and rows[0]["attending"] == 0
assert rows[0]["has_allergy"] is None and rows[0]["allergy_notes"] is None and rows[0]["submission_count"] == 2
# Validation errors keep the typed values.
r = submit({"full_name": "Bob", "email": "bob@x.com", "attending": "yes", "has_allergy": "yes",
            "allergy_notes": ""})
assert r.status_code == 422 and "Please tell us what you are allergic to." in r.get_data(as_text=True)
assert 'value="bob@x.com"' in r.get_data(as_text=True)
r = submit({"full_name": "Bob", "email": "bob@x.com", "attending": "yes", "has_allergy": "yes",
            "allergy_notes": "z" * 201})
assert r.status_code == 422 and "at most 200 characters" in r.get_data(as_text=True)
r = submit({"full_name": "Bob", "email": "not-an-email", "attending": ""})
assert r.status_code == 422 and "valid email" in r.get_data(as_text=True) and "Please choose Yes or No." in r.get_data(as_text=True)
# Honeypot: looks like success, nothing stored.
r = anon.post(f"/i/{SLUG}", data={"csrf": PUB_CSRF, "website": "http://spam", "full_name": "Bot",
                                  "email": "bot@spam.com", "attending": "yes", "has_allergy": "no"},
              headers=new_ip())
assert r.status_code == 303 and len(INV.STORE.responses(EID)) == 1
# CSRF.
r = submit({"full_name": "Bob", "email": "bob@x.com", "attending": "no"}, token="0" * 64)
assert r.status_code == 400 and "expired" in r.get_data(as_text=True)
# Oversized body.
r = anon.post(f"/i/{SLUG}", data={"csrf": PUB_CSRF, "full_name": "x" * 70000}, headers=new_ip())
assert r.status_code == 413
for name, email, att, alg, notes in (("Bob Tan", "bob@x.com", "yes", "no", ""),
                                     ("Chi Pham", "chi@y.vn", "yes", "yes", "Shellfish, gluten"),
                                     ("Dan Ho", "dan@z.my", "no", "", ""),
                                     ("=cmd()", "eve@x.com", "yes", "no", "")):
    assert submit({"full_name": name, "email": email, "attending": att, "has_allergy": alg,
                   "allergy_notes": notes}).status_code == 303, name
print("ok public form")

# ------------------------------------------------------------------ rate limits
ip = new_ip()
for i in range(INV.IP_LIMIT):
    r = submit({"full_name": "", "email": "", "attending": ""}, ip=ip)
    assert r.status_code == 422, (i, r.status_code)
r = submit({"full_name": "Late", "email": "late@x.com", "attending": "no"}, ip=ip)
assert r.status_code == 429 and "Too many submissions" in r.get_data(as_text=True)
for i in range(INV.EMAIL_LIMIT - 1):                 # bob already submitted once
    assert submit({"full_name": "Bob Tan", "email": "BOB@x.com", "attending": "yes",
                   "has_allergy": "no"}).status_code == 303, i
r = submit({"full_name": "Bob Tan", "email": "bob@x.com", "attending": "yes", "has_allergy": "no"})
assert r.status_code == 429 and "too many times" in r.get_data(as_text=True)
assert INV.STORE.hits("ip:10.0.0.1", 3600) <= 1     # other IPs are not counted together
print("ok rate limits")

# ------------------------------------------------------------------ admin view, search, export
counts = INV.STORE.counts(EID)
assert counts == {"total": 5, "attending": 3, "not_attending": 2, "allergies": 1}, counts
page = sales.get(f"/invitation/{EID}").get_data(as_text=True)
for label, n in (("Total", 5), ("Attending", 3), ("Not attending", 2), ("With allergies", 1)):
    assert re.search(rf"<dt>{label}</dt><dd>{n}</dd>", page), label
assert "edited ×2" in page and "Shellfish, gluten" in page
page = sales.get(f"/invitation/{EID}?q=CHI").get_data(as_text=True)
assert "chi@y.vn" in page and "bob@x.com" not in page
page = sales.get(f"/invitation/{EID}?filter=allergy").get_data(as_text=True)
assert "Chi Pham" in page and "Dan Ho" not in page
page = sales.get(f"/invitation/{EID}?q=%25").get_data(as_text=True)    # LIKE wildcard escaped
assert "No responses match." in page
r = sales.get(f"/invitation/{EID}/export?format=csv")
assert r.status_code == 200 and r.headers["Cache-Control"] == "no-store"
assert re.search(r'filename="dpm-invitation-dpm-gala-dinner-2027-\d{8}\.csv"', r.headers["Content-Disposition"])
text = r.get_data().decode("utf-8")
assert text.startswith("\ufeff")
rows = list(csv.reader(io.StringIO(text[1:])))
assert rows[0] == export.HEADERS and len(rows) == 6, rows
chi = next(x for x in rows if x[1] == "chi@y.vn")
assert chi[2:5] == ["Yes", "Yes", "Shellfish, gluten"], chi
dan = next(x for x in rows if x[1] == "dan@z.my")
assert dan[2:5] == ["No", "", ""], dan
assert any(x[0] == "'=cmd()" for x in rows)
r = sales.get(f"/invitation/{EID}/export?format=xlsx&filter=attending")
assert r.status_code == 200 and r.mimetype.endswith("spreadsheetml.sheet")
ws = load_workbook(io.BytesIO(r.get_data())).active
vals = [[c.value for c in row] for row in ws.iter_rows()]
assert vals[0] == export.HEADERS and len(vals) == 4, vals            # 3 attending
assert sales.get(f"/invitation/{EID}/export?format=pdf").status_code == 400
assert kvb.get(f"/invitation/{EID}/export?format=csv", headers={"Accept": "*/*"}).status_code == 403
assert anon.get(f"/invitation/{EID}/export?format=csv").status_code == 401
print("ok admin view + export")

# ------------------------------------------------------------------ deletes
SALES_TOKEN = csrf(sales, f"/invitation/{EID}")
dan_row = next(x for x in INV.STORE.responses(EID) if x["email"] == "dan@z.my")
r = sales.post(f"/invitation/{EID}/responses/{dan_row['id']}/delete", data={"csrf_token": SALES_TOKEN})
assert r.status_code == 303 and INV.STORE.counts(EID)["total"] == 4
assert sales.post(f"/invitation/{EID}/responses/{dan_row['id']}/delete",
                  data={"csrf_token": SALES_TOKEN}).status_code == 404
assert sales.post(f"/invitation/{EID}/responses/1/delete", data={}).status_code == 400   # no CSRF
assert sales.post(f"/invitation/{EID}/delete", data={"csrf_token": SALES_TOKEN, "confirm": SLUG}).status_code == 403
print("ok response delete")

# ------------------------------------------------------------------ non-admin DPM user manages events
r = sales.post("/invitation/new", data={
    "csrf_token": SALES_TOKEN, "name": "Sales Team Lunch", "slug": "sales-lunch", "closes_at": "",
    "image": (io.BytesIO(img_bytes()), "lunch.jpg")}, content_type="multipart/form-data")
assert r.status_code == 303, r.status_code
SID = int(r.headers["Location"].rstrip("/").split("/")[-1])
sev = INV.STORE.event(SID)
assert sev["created_by"] == "sales" and sev["status"] == "open"
r = sales.post(f"/invitation/{SID}/edit", data={
    "csrf_token": SALES_TOKEN, "name": "Sales Team Lunch 2027", "slug": "sales-lunch-2027",
    "closes_at": "2030-06-01T12:00",
    "image": (io.BytesIO(img_bytes("PNG", size=(800, 1200))), "new.png")}, content_type="multipart/form-data")
assert r.status_code == 303, r.status_code
sev2 = INV.STORE.event(SID)
assert sev2["name"] == "Sales Team Lunch 2027" and sev2["slug"] == "sales-lunch-2027" and sev2["updated_by"] == "sales"
assert sev2["image_key"] != sev["image_key"] and not (INV.SETTINGS.files / sev["image_key"]).exists()
assert rules.format_wib(sev2["closes_at"]) == "01 Jun 2030, 12:00 WIB"
r = sales.post(f"/invitation/{SID}/status", data={"csrf_token": SALES_TOKEN, "status": "closed"})
assert r.status_code == 303 and INV.STORE.event(SID)["status"] == "closed"
assert "This invitation is closed" in anon.get("/i/sales-lunch-2027").get_data(as_text=True)
r = sales.post(f"/invitation/{SID}/status", data={"csrf_token": SALES_TOKEN, "status": "open"})
assert r.status_code == 303 and INV.STORE.event(SID)["status"] == "open"
assert sales.get(f"/invitation/{SID}/export?format=xlsx").status_code == 200
assert sales.post(f"/invitation/{SID}/status", data={"csrf_token": "bad", "status": "closed"}).status_code == 400
# Deleting a whole event stays admin-only, even for the user who created it.
r = sales.post(f"/invitation/{SID}/delete", data={"csrf_token": SALES_TOKEN, "confirm": "sales-lunch-2027"})
assert r.status_code == 403 and INV.STORE.event(SID) is not None
# Users without Invitation access are refused on every write.
for path, data in ((f"/invitation/{SID}/edit", {"name": "x"}), (f"/invitation/{SID}/status", {"status": "closed"}),
                   (f"/invitation/{SID}/delete", {"confirm": "sales-lunch-2027"})):
    assert kvb.post(path, data=data, headers={"Accept": "*/*"}).status_code == 403, path
    assert anon.post(path, data=data).status_code == 401, path
r = admin.post(f"/invitation/{SID}/delete", data={"csrf_token": ADMIN_TOKEN, "confirm": "sales-lunch-2027"})
assert r.status_code == 303 and INV.STORE.event(SID) is None
print("ok non-admin DPM user manages events")

# ------------------------------------------------------------------ close / auto-close / edit
r = admin.post(f"/invitation/{EID}/status", data={"csrf_token": ADMIN_TOKEN, "status": "closed"})
assert r.status_code == 303 and INV.STORE.event(EID)["status"] == "closed"
r = anon.get(f"/i/{SLUG}")
assert r.status_code == 200 and "This invitation is closed" in r.get_data(as_text=True)
assert 'name="full_name"' not in r.get_data(as_text=True)
r = submit({"full_name": "Late Guy", "email": "late@x.com", "attending": "no"})
assert r.status_code == 403 and INV.STORE.counts(EID)["total"] == 4
assert "Change my answer" not in anon.get(f"/i/{SLUG}/thanks").get_data(as_text=True)
admin.post(f"/invitation/{EID}/status", data={"csrf_token": ADMIN_TOKEN, "status": "open"})
assert "Confirm your attendance" in anon.get(f"/i/{SLUG}").get_data(as_text=True)
INV.STORE.update_event(EID, "test", closes_at=time.time() - 5)        # the closing time passed
assert "This invitation is closed" in anon.get(f"/i/{SLUG}").get_data(as_text=True)
assert submit({"full_name": "Late Guy", "email": "late@x.com", "attending": "no"}).status_code == 403
page = admin.get(f"/invitation/{EID}").get_data(as_text=True)
assert "Closed automatically" in page and "Clear or move the closing time to reopen." in page
old_key = INV.STORE.event(EID)["image_key"]
r = admin.post(f"/invitation/{EID}/edit", data={
    "csrf_token": ADMIN_TOKEN, "name": "DPM Gala Dinner 2027 Hanoi", "slug": "gala-2027",
    "closes_at": "", "clear_closes_at": "1",
    "image": (io.BytesIO(img_bytes("PNG", size=(800, 1200), color=(200, 10, 10))), "new.png")},
    content_type="multipart/form-data")
assert r.status_code == 303, r.status_code
ev = INV.STORE.event(EID)
assert ev["slug"] == "gala-2027" and ev["closes_at"] is None and ev["image_key"] != old_key
assert not (INV.SETTINGS.files / old_key).exists() and (INV.SETTINGS.files / ev["image_key"]).exists()
assert anon.get(f"/i/{SLUG}").status_code == 404 and anon.get("/i/gala-2027").status_code == 200
r = admin.post(f"/invitation/{EID}/edit", data={"csrf_token": ADMIN_TOKEN, "name": "Gala",
                                                "slug": "gala-2027", "closes_at": "2020-01-01T10:00"})
assert r.status_code == 422 and "future" in r.get_data(as_text=True)
print("ok close / auto-close / edit")

# ------------------------------------------------------------------ event delete
r = admin.post(f"/invitation/{EID}/delete", data={"csrf_token": ADMIN_TOKEN, "confirm": "wrong"})
assert r.status_code == 303 and INV.STORE.event(EID) is not None
r = admin.post(f"/invitation/{EID}/delete", data={"csrf_token": ADMIN_TOKEN, "confirm": "gala-2027"})
assert r.status_code == 303 and INV.STORE.event(EID) is None
assert INV.STORE.responses(EID) == [] and not (INV.SETTINGS.files / ev["image_key"]).exists()
assert anon.get("/i/gala-2027").status_code == 404
print("ok event delete")

# The schema is additive: opening an existing DB again changes nothing.
from invitation.store import Store  # noqa: E402
again = Store(INV.SETTINGS.db)
assert again.events() == INV.STORE.events()
T.assert_isolated(TEST_ROOT)
print("ALL OK test_invitation")
