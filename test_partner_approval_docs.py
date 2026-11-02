#!/usr/bin/env python3
"""Partner Approval (Beta): form rules, summary, agreement fill, signatures.

python3 test_partner_approval_docs.py   (no LibreOffice needed)
"""
import base64
import io
import os
import shutil
import tempfile
from pathlib import Path

import testkit as T

TEST_ROOT = T.hermetic_env()

from partner_approval import docs, forms  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="pa-docs-"))
CM = Path(__file__).resolve().parent / "partner_approval" / "agreement_templates" / "CM.docx"

AGREEMENT = {
    "request_type": "agreement", "contract_type": "NEW", "sales_team": "Shawn Team",
    "sales_uid": "100200", "role": "CM", "uid": "999001", "partner_name": "Test Partner",
    "passport": "X0000001", "address": "1 Sample Road\nSample City",
    "team_upline": "ISS L\u00ca TH\u00c0NH T\u00c2M 446157 > RM 543936 LODIS",
    "background": "Dummy background text.", "allowance_scheme": "Net Deposit : 12%",
    "remarks": "Rebate will be set by superior", "region": "Vietnam",
    "effective_date": "2026-11-01", "onboard_date": "2026-10-09",
    "submission_date": "2026-10-30", "partner_email": "partner@example.com",
    "desk": "VN Desk", "create_agreement": "Yes",
}
DOWNGRADE = {
    "request_type": "downgrade", "role": "RM", "uid": "999002", "partner_name": "Test Two",
    "effective_date": "2026-11-01", "new_position": "MIB",
    "downgrade_type": "Downgrade CRM Role", "desk": "ASIA Desk",
}

# ------------------------------------------------------------ form rules
data, err = forms.validate(AGREEMENT)
assert not err, err
assert data["team_upline"] == AGREEMENT["team_upline"]          # unicode kept as typed
assert data["address"] == "1 Sample Road\nSample City"            # long text keeps lines

data, err = forms.validate({**AGREEMENT, "request_type": "special_mib"})
assert not err, err

# every required agreement field is required
for name in ("contract_type", "sales_team", "sales_uid", "role", "uid", "partner_name",
             "passport", "address", "team_upline", "background", "allowance_scheme",
             "region", "effective_date", "onboard_date", "submission_date",
             "partner_email", "desk"):
    _d, err = forms.validate({**AGREEMENT, name: ""})
    assert name in err, (name, err)
for name in ("remarks", "create_agreement"):                     # optional
    _d, err = forms.validate({**AGREEMENT, name: ""})
    assert not err, (name, err)

_d, err = forms.validate({**AGREEMENT, "sales_uid": "12a"})
assert "sales_uid" in err
_d, err = forms.validate({**AGREEMENT, "uid": "0"})
assert "uid" in err
_d, err = forms.validate({**AGREEMENT, "role": "CEO"})
assert "role" in err
_d, err = forms.validate({**AGREEMENT, "region": "Mars"})
assert "region" in err
_d, err = forms.validate({**AGREEMENT, "partner_email": "not-an-email"})
assert "partner_email" in err
_d, err = forms.validate({**AGREEMENT, "onboard_date": "31/12/2026"})
assert "onboard_date" in err
_d, err = forms.validate({**AGREEMENT, "background": "x" * 5001})
assert "background" in err
_d, err = forms.validate({"request_type": "upgrade"})
assert "request_type" in err

# Downgrade / Exit: a different, smaller field set
for rtype in ("downgrade", "exit"):
    data, err = forms.validate({**DOWNGRADE, "request_type": rtype})
    assert not err, (rtype, err)
    for name in ("role", "uid", "partner_name", "effective_date", "new_position",
                 "downgrade_type", "desk"):
        _d, err = forms.validate({**DOWNGRADE, "request_type": rtype, name: ""})
        assert name in err, (rtype, name, err)
    # email optional here, but still checked when given
    assert not forms.validate({**DOWNGRADE, "request_type": rtype})[1]
    assert "partner_email" in forms.validate({**DOWNGRADE, "request_type": rtype,
                                              "partner_email": "nope"})[1]
# agreement-only fields are dropped from a downgrade, even when posted
data, err = forms.validate({**DOWNGRADE, "passport": "SNEAKY", "allowance_scheme": "99%"})
assert "passport" not in data and "allowance_scheme" not in data
# and downgrade-only fields are not asked for an agreement
data, err = forms.validate({**AGREEMENT, "new_position": "Client"})
assert "new_position" not in data
assert forms.REQUEST_TYPES["downgrade"] == "Dowgrade"            # the source form's spelling

# control characters are removed, single-line fields collapsed
data, _ = forms.validate({**AGREEMENT, "partner_name": "  Test \x00\x07 Partner \n X "})
assert data["partner_name"] == "Test Partner X", repr(data["partner_name"])

# ------------------------------------------------------------- summary text
data, _ = forms.validate(AGREEMENT)
s = forms.summary(data)
expected_head = """Partnership Agreement
Effective Date: 11/01/2026
Onboard Date: 10/09/2026
Team: Shawn
Region: Vietnam
Team Upline:
ISS L\u00ca TH\u00c0NH T\u00c2M 446157 > RM 543936 LODIS
Background:
Dummy background text.
Agreement Type: New Join
Allowance Scheme:
Net Deposit : 12%
Remarks:
Rebate will be set by superior"""
assert s.startswith(expected_head), s
s = forms.summary(forms.validate(DOWNGRADE)[0])
assert s.startswith("Dowgrade\nEffective Date: 11/01/2026"), s
assert "New Position: MIB" in s and "Type of Downgrade: Downgrade CRM Role" in s

# ------------------------------------------------------------ templates
assert docs.template_for(CM.parent, "CM") == CM
for role in ("BM", "RM", "MIB", "NCIB", "RD"):
    assert docs.template_for(CM.parent, role) is None, role
assert docs.template_for(CM.parent, "../CM") is None

# The committed template is blank and has no staff names in its metadata.
from docx import Document  # noqa: E402
tpl = Document(str(CM))
assert tpl.core_properties.author == "Dupoin Markets Ltd"
assert tpl.core_properties.last_modified_by == "Dupoin Markets Ltd"
assert set(docs.unfilled_placeholders(CM, signed=True)) >= {
    "made_on", "region", "partner_name", "passport", "address", "uid",
    "primary_location", "as_of", "sign_partner_name", "sign_passport", "signature"}

# ------------------------------------------------------------- fill (unsigned)
data, _ = forms.validate(AGREEMENT)
out = TMP / "a.docx"
filled = docs.fill_agreement(CM, data, out)
assert docs.unfilled_placeholders(out) == [], docs.unfilled_placeholders(out)
assert docs.company_column_blank(out)
d = Document(str(out))
text = "\n".join(p.text for p in d.paragraphs)
assert "This Agreement is made on 1 November 2026" in text
assert "Region/Country\t\t: Vietnam" in text
assert "Name\t\t\t\t\t: Test Partner" in text
assert "Passport/National ID \t: X0000001" in text
assert "Address\t\t\t\t\t: 1 Sample Road, Sample City" in text
assert "The primary location is Vietnam, with" in text
assert "as of 1 November 2026, with each Party" in text
assert "Remark: Rebate will be set by superior" in text
assert "Agreed Allowance Scheme:\nNet Deposit : 12%" in text
# Schedule A template content is untouched
for keep in ("Net Deposit Bonus : 2%", "Rebates:", "Forex: 6.6", "Notes"):
    assert keep in text, keep
tables = ["|".join(c.text for c in r.cells) for t in d.tables for r in t.rows]
assert "Dupoin UID|999001" in tables, tables
assert "Below $10,000|$0" in tables and "$60,001 and Above|$1,500" in tables
# sign date is left empty until the partner signs
assert "Date\t\t\t: " in text and "Date\t\t\t: 2" not in text
assert d.core_properties.author == "Dupoin Markets Ltd"

# ---------------------------------------------------------- signatures
typed = docs.typed_signature("L\u00ea Th\u00e0nh T\u00e2m")       # Vietnamese in the script font
assert typed.startswith(docs.PNG_MAGIC)
from PIL import Image, ImageDraw  # noqa: E402
assert Image.open(io.BytesIO(typed)).width > 100
for bad in ("", " ", "x", "y" * 81):
    try:
        docs.typed_signature(bad)
        raise AssertionError(f"typed {bad!r} accepted")
    except docs.SignatureError:
        pass


def canvas_png(draw=True, background=None):
    img = Image.new("RGBA", (600, 200), background or (0, 0, 0, 0))
    if draw:
        ImageDraw.Draw(img).line([(40, 150), (200, 40), (330, 160), (560, 60)],
                                 fill=(16, 32, 96, 255), width=4)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


drawn = docs.drawn_signature(canvas_png())
im = Image.open(io.BytesIO(drawn))
assert im.width < 600 and im.height < 200, im.size              # cropped to the ink
docs.drawn_signature(canvas_png(background=(255, 255, 255, 255)))  # white canvas is fine
for bad in ("", "data:image/jpeg;base64,AAAA", canvas_png(draw=False),
            canvas_png(draw=False, background=(255, 255, 255, 255)),
            "data:image/png;base64," + base64.b64encode(b"GIF89a....").decode(),
            "data:image/png;base64,!!!notbase64",
            "data:image/png;base64," + "A" * (3 * 1024 * 1024)):
    try:
        docs.drawn_signature(bad)
        raise AssertionError(f"drawn {bad[:40]!r} accepted")
    except docs.SignatureError:
        pass

# ----------------------------------------------------------- fill (signed)
out2 = TMP / "signed.docx"
filled = docs.fill_agreement(CM, data, out2, sign_date="2026-11-02", signature_png=drawn)
assert "signature" in filled
assert docs.unfilled_placeholders(out2, signed=True) == []
assert docs.company_column_blank(out2), "company signature column must stay blank"
text = "\n".join(p.text for p in Document(str(out2)).paragraphs)
assert "Date\t\t\t: 2 November 2026" in text
assert "Passport/ID No\t: X0000001" in text

# ------------------------------------------------------ uploads by magic bytes
assert docs.sniff_upload(b"%PDF-1.7\n") == "pdf"
assert docs.sniff_upload(docs.PNG_MAGIC + b"....") == "png"
assert docs.sniff_upload(b"\xff\xd8\xff\xe0..") == "jpg"
for bad in (b"MZ\x90\x00", b"<html>", b"PK\x03\x04", b""):
    assert docs.sniff_upload(bad) is None, bad

# --------------------------------------------- PDF conversion without soffice
try:
    docs.docx_to_pdf(out, TMP / "x.pdf", soffice=str(TMP / "no-such-soffice"))
    raise AssertionError("conversion without LibreOffice must raise")
except docs.ConversionUnavailable as exc:
    assert "libreoffice-writer" in str(exc)
assert not (TMP / "x.pdf").exists()

# with soffice (only where it is installed: the box yes, the server maybe not)
if docs.find_soffice():
    pdf = docs.docx_to_pdf(out2, TMP / "signed.pdf")
    assert pdf.read_bytes().startswith(b"%PDF-")
    print("  (real LibreOffice conversion checked)")
else:
    print("  (LibreOffice not installed here: real conversion skipped)")

shutil.rmtree(TMP, True)
print("test_partner_approval_docs: ok")
