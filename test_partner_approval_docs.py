#!/usr/bin/env python3
"""Partner Approval (Beta): form rules, summary, agreement fill, signatures.

python3 test_partner_approval_docs.py   (no LibreOffice needed)
"""
import base64
import io
import os
import re
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
    "team_upline": "446157 - 543936",
    "background": "Dummy background text.", "allowance_scheme": "Net Deposit : 12%",
    "remarks": "Rebate will be set by superior", "region": "Vietnam",
    "effective_date": "2026-11-01", "onboard_date": "2026-10-09",
    "submission_date": "2026-10-30", "partner_email": "partner@example.com",
    "desk": "VN Desk",
}
DOWNGRADE = {
    "request_type": "downgrade", "role": "RM", "uid": "999002", "partner_name": "Test Two",
    "effective_date": "2026-11-01", "new_position": "MIB",
    "downgrade_type": "Downgrade CRM Role", "desk": "ASIA Desk",
}

# ------------------------------------------------------------ form rules
data, err = forms.validate(AGREEMENT)
assert not err, err
assert data["team_upline"] == "446157 - 543936"
assert data["address"] == "1 Sample Road\nSample City"            # long text keeps lines

# Special Terms (was "Special MIB"): MIB role only, server side
data, err = forms.validate({**AGREEMENT, "request_type": "special_terms", "role": "MIB"})
assert not err, err
assert forms.REQUEST_TYPES["special_terms"] == "Special Terms"
assert "special_mib" not in forms.REQUEST_TYPES
for role in ("CM", "BM", "RM", "NCIB", "RD"):
    _d, err = forms.validate({**AGREEMENT, "request_type": "special_terms", "role": role})
    assert err.get("role") == "Special Terms requests are for the MIB role only.", (role, err)
for role in forms.ROLES:                                          # Agreement: every role
    assert not forms.validate({**AGREEMENT, "role": role})[1], role
assert forms.roles_for("agreement") == ("CM", "BM", "RM", "MIB", "NCIB", "RD")
assert forms.roles_for("special_terms") == ("MIB",)
# a form opened before the rename still posts "special_mib"
data, err = forms.validate({**AGREEMENT, "request_type": "special_mib", "role": "MIB"})
assert not err and data["request_type"] == "special_terms", (data, err)
assert "Partnership Agreement (Special Terms)" in forms.summary(data)

# "Do u want to create agreement?" is gone: not a field, not in the summary
assert "create_agreement" not in {f["name"] for f in forms.FIELDS}
data, err = forms.validate({**AGREEMENT, "create_agreement": "No"})
assert not err and "create_agreement" not in data
s_old = forms.summary({**forms.validate(AGREEMENT)[0], "create_agreement": "No"})  # old stored data
assert "reate agreement" not in s_old
assert all("reate agreement" not in r[0] for _sec, rows in forms.details(
    {**forms.validate(AGREEMENT)[0], "create_agreement": "Yes"}) for r in rows)

# Team Upline: 6-digit groups joined by " - "
for raw, want in (("123456", "123456"), ("123456123456", "123456 - 123456"),
                  ("123456 - 123456 - 123456", "123456 - 123456 - 123456"),
                  ("123456-123456", "123456 - 123456"), (" 123456  -  654321 ", "123456 - 654321"),
                  ("123456\u2013654321", "123456 - 654321")):
    data, err = forms.validate({**AGREEMENT, "team_upline": raw})
    assert not err and data["team_upline"] == want, (raw, data.get("team_upline"), err)
    assert forms.UPLINE_RE.match(data["team_upline"])
for raw, part in (("12345", "has 5"), ("123456 - 1234", "has 4"), ("1234567", "has 1"),
                  ("123456 > RM 543936", "digits only"), ("ISS 446157", "digits only"),
                  ("123456 - ", None), (" - ", "digits only")):
    _d, err = forms.validate({**AGREEMENT, "team_upline": raw})
    if part is None:                       # trailing dash only: still a valid group
        assert not err and _d["team_upline"] == "123456", (raw, err)
        continue
    assert part in err.get("team_upline", ""), (raw, err)
assert "Team Upline:\n446157 - 543936" in forms.summary(forms.validate(AGREEMENT)[0])

# every required agreement field is required
for name in ("contract_type", "sales_team", "sales_uid", "role", "uid", "partner_name",
             "passport", "address", "team_upline", "background", "allowance_scheme",
             "region", "effective_date", "onboard_date", "submission_date",
             "partner_email", "desk"):
    _d, err = forms.validate({**AGREEMENT, name: ""})
    assert name in err, (name, err)
for name in ("remarks",):                                        # optional
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
446157 - 543936
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
for role in ("BM", "RM"):                                          # v1.35
    assert docs.template_for(CM.parent, role) == CM.parent / f"{role}.docx", role
assert docs.template_for(CM.parent, "PIF") is None                # PIF is not an agreement

# -------------------------- roles without their own file borrow one (v1.37)
import shutil as _sh  # noqa: E402
import tempfile as _tf  # noqa: E402
EXPECT_FALLBACK = {"MIB": "RM", "NCIB": "RM", "RD": "CM"}
assert docs.TEMPLATE_FALLBACK == EXPECT_FALLBACK
for role, src in EXPECT_FALLBACK.items():
    assert docs.template_info(CM.parent, role) == (CM.parent / f"{src}.docx", src), role
    assert docs.template_for(CM.parent, role) == CM.parent / f"{src}.docx", role
for role in ("CM", "BM", "RM"):                                     # own files, not borrowed
    assert docs.template_info(CM.parent, role)[1] is None, role
assert docs.template_info(CM.parent, "PIF") == (None, None)
assert docs.template_info(CM.parent, "XYZ") == (None, None)
# a real MIB.docx / NCIB.docx / RD.docx wins as soon as it is in the folder
tmpl = Path(_tf.mkdtemp())
for n in ("CM", "BM", "RM"):
    _sh.copy(CM.parent / f"{n}.docx", tmpl / f"{n}.docx")
_sh.copy(CM.parent / "BM.docx", tmpl / "MIB.docx")               # stand-in "real" MIB file
assert docs.template_info(tmpl, "MIB") == (tmpl / "MIB.docx", None)
assert docs.template_info(tmpl, "NCIB") == (tmpl / "RM.docx", "RM")
assert docs.role_swaps(tmpl / "MIB.docx", "MIB") == []              # own file: no rewording
# nothing to borrow -> None, as before
(tmpl / "RM.docx").unlink()
assert docs.template_info(tmpl, "NCIB") == (None, None)

ROLE_DATA = {"partner_name": "Dummy Partner", "region": "Testland", "passport": "X0000000",
             "address": "1 Example Street", "uid": "999999", "effective_date": "2026-11-01"}


def _filled_text(role, template=None):
    template = template or docs.template_for(CM.parent, role)
    out = Path(_tf.mkdtemp()) / "a.docx"
    filled = docs.fill_agreement(template, {**ROLE_DATA, "role": role}, out)
    d = Document(str(out))
    return filled, "\n".join(p_.text for p_ in docs._all_paragraphs(d)), d


from docx import Document  # noqa: E402
for role, title in (("MIB", "Master IB"), ("NCIB", "NCIB")):
    filled, text, d = _filled_text(role)
    assert "role_wording" in filled and "uid" in filled and "partner_name" in filled, filled
    assert "Relationship Manager" not in text and not re.search(r"\bRM\b", text), role
    assert f"referred to as “{title}”" in text and f"SIGNED by the {title}" in text, role
    assert f"an {role}'s customer" in text and f"disbursed to the {role} prior" in text, role
    assert "Dummy Partner" in text and "999999" in text
    heads = [c.text.strip() for c in docs._signature_table(d)[0].rows[0].cells]
    assert heads[0] == "SIGNED by the Company" and heads[-1] == f"SIGNED by the {title}", heads
filled, text, d = _filled_text("RD")
assert "Country Manager" not in text and not re.search(r"\bCM\b", text)
assert "referred to as “Regional Director”" in text and "SIGNED by the Regional Director" in text
assert "The Regional Director will report directly to the Director of Business Development." in text
assert "report directly to the Regional Director" not in text
assert "disbursed to the RD prior" in text
# subordinate wording in CM.docx is about other roles and stays as written
assert "Support Relationship Managers and Master IBs" in text and "the part of the RM/BM" in text
# the real roles keep their own template wording untouched
for role, title in (("CM", "Country Manager"), ("BM", "Business Manager"), ("RM", "Relationship Manager")):
    filled, text, _ = _filled_text(role)
    assert "role_wording" not in filled and f"referred to as “{title}”" in text, role
# a label split over several runs is still replaced, keeping the first run's format
doc = Document()
para = doc.add_paragraph()
for chunk in ("the Relat", "ionship Man", "ager is"):
    para.add_run(chunk).bold = True
assert docs.swap_role_wording(doc, [("Relationship Manager", "Master IB")]) == 1
assert para.text == "the Master IB is" and all(r.bold for r in para.runs if r.text)
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
typed = docs.typed_signature("Andrew Tan")                       # default script font
assert typed.startswith(docs.PNG_MAGIC)
from PIL import Image, ImageDraw  # noqa: E402
assert Image.open(io.BytesIO(typed)).width > 100
for bad in ("", " ", "x", "y" * 81):
    try:
        docs.typed_signature(bad)
        raise AssertionError(f"typed {bad!r} accepted")
    except docs.SignatureError:
        pass


# ------------------------------------------------- pluggable signature fonts
fonts = docs.signature_fonts()
assert 1 <= len(fonts) <= docs.MAX_SIGNATURE_FONTS, fonts
assert [f["id"] for f in fonts] == sorted(f["id"] for f in fonts)
VI_FONTS = ["4-GreatVibes-Regular", "5-DancingScript-Regular", "6-Allura-Regular"]
assert [f["id"] for f in fonts] == ["1-DarceyOliver-Regular", "2-Budaphest-Regular", "3-Augustia"] + VI_FONTS
assert [f["label"] for f in fonts] == ["Darcey Oliver", "Budaphest", "Augustia",
                                       "Great Vibes", "Dancing Script", "Allura"]
assert docs.MAX_SIGNATURE_FONTS >= 6
for fid in VI_FONTS:                            # each free font ships its OFL licence
    stem = fid.split("-", 1)[1].replace("-Regular", "")
    assert "SIL Open Font License" in (docs.SIGNATURE_FONTS / f"{stem}-OFL.txt").read_text()
assert docs.SCRIPT_FONT.is_file() and docs.SCRIPT_FONT.name == fonts[0]["file"]
assert all(f["label"] and "Regular" not in f["label"] for f in fonts), fonts
assert docs._font_label("2-MyBrandScript-Regular") == "My Brand Script"
assert docs._font_label("DancingScript-VariableFont_wght") == "Dancing Script"
# glyph coverage: a style that lacks a letter of the name is refused, never drawn as tofu
try:
    from fontTools.ttLib import TTFont  # noqa: E402  (optional, dev-only cross-check)
except ImportError:
    TTFont = None
for f in fonts:
    if TTFont:
        cmap = {chr(c) for c in TTFont(docs.SIGNATURE_FONTS / f["file"]).getBestCmap()}
        # measured coverage = the cmap, plus letters PIL composes from base + combining marks
        assert cmap & docs.SIGNATURE_CHARSET <= set(f["chars"]), f["id"]
        if f["id"] in ("1-DarceyOliver-Regular", "2-Budaphest-Regular", "3-Augustia"):
            assert set(f["chars"]) == cmap & docs.SIGNATURE_CHARSET, f["id"]
    assert set("ABCXYZabcxyz") <= set(f["chars"]), f["id"]
    assert docs.typed_signature("Andrew Tan", f["id"]).startswith(docs.PNG_MAGIC)
VI = "Nguy\u1ec5n V\u0103n \u0110\u1ee9c"
assert docs.unsupported_chars(VI, "1-DarceyOliver-Regular") == ["\u1ec5", "\u0110", "\u1ee9"]
assert docs.unsupported_chars("Jean-Luc O'Brien 2", "2-Budaphest-Regular") == ["-", "'", "2"]
assert docs.unsupported_chars("Jean-Luc O'Brien 2", "1-DarceyOliver-Regular") == []
assert docs.unsupported_chars("Nguye\u0302\u0303n", "1-DarceyOliver-Regular") == ["\u1ec5"]  # NFD -> NFC
assert docs.unsupported_chars("Andrew \U0001F600", "3-Augustia") == ["\U0001F600"]
assert docs.first_font_for("Jean-Luc") == "1-DarceyOliver-Regular"
assert docs.first_font_for(VI) == "4-GreatVibes-Regular"
assert docs.first_font_for("Lê Thành Tâm") == "4-GreatVibes-Regular"
assert docs.first_font_for("\u738b\u5c0f\u660e") is None              # not Latin/Vietnamese -> draw
# every Vietnamese letter is drawable in at least the 3 OFL styles
ALL_VI = ("ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúýĂăĐđĨĩŨũƠơƯưẠạẢảẤấẦầẨẩẪẫẬậẮắẰằẲẳẴẵẶặẸẹẺẻẼẽẾếỀềỂểỄễỆệ"
          "ỈỉỊịỌọỎỏỐốỒồỔổỖỗỘộỚớỜờỞởỠỡỢợỤụỦủỨứỪừỬửỮữỰựỲỳỴỵỶỷỸỹ")
for fid in VI_FONTS:
    assert docs.unsupported_chars(ALL_VI, fid) == [], fid
    assert docs.typed_signature(VI, fid).startswith(docs.PNG_MAGIC)
for f in fonts[:3]:                              # Bayu's fonts: refused, never tofu
    try:
        docs.typed_signature(VI, f["id"])
        raise AssertionError(f"tofu accepted in {f['id']}")
    except docs.SignatureError as exc:
        assert "cannot show some letters" in str(exc)
assert sum(not docs.unsupported_chars(VI, f["id"]) for f in fonts) >= 3
try:
    docs.typed_signature("O'Brien", "2-Budaphest-Regular")
    raise AssertionError("Budaphest has no apostrophe")
except docs.SignatureError:
    pass
assert docs.signature_font_path(None) == docs.signature_font_path("") == docs.SIGNATURE_FONTS / fonts[0]["file"]
for bad in ("../signature_fonts/3-Augustia", "../../app", "/etc/passwd", "3-Augustia.ttf",
            "3-augustia", "nope", "3-Augustia\x00", "..", "x" * 90):
    try:
        docs.signature_font_path(bad)
        raise AssertionError(f"font {bad!r} accepted")
    except docs.SignatureError:
        pass
import tempfile, shutil as _sh  # noqa: E402
with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    src = docs.SIGNATURE_FONTS / "3-Augustia.ttf"
    _sh.copy(src, tmp / "1-Zeta-Regular.ttf"); _sh.copy(src, tmp / "2-Alpha.otf")
    _sh.copy(src, tmp / "bad name.ttf"); _sh.copy(src, tmp / ".hidden.ttf"); _sh.copy(src, tmp / "notes.txt")
    outside = Path(tempfile.mkdtemp()); _sh.copy(src, outside / "Evil.ttf")
    (tmp / "Linked.ttf").symlink_to(outside / "Evil.ttf")
    got = docs.signature_fonts(tmp)
    assert [f["id"] for f in got] == ["1-Zeta-Regular", "2-Alpha"], got
    assert [f["label"] for f in got] == ["Zeta", "Alpha"], got
    assert docs.signature_font_path("2-Alpha", tmp) == tmp / "2-Alpha.otf"
    for bad in ("Linked", "bad name", ".hidden", "notes"):
        try:
            docs.signature_font_path(bad, tmp)
            raise AssertionError(bad)
        except docs.SignatureError:
            pass
    for i in range(10):
        _sh.copy(src, tmp / f"F{i}.ttf")
    assert len(docs.signature_fonts(tmp)) == docs.MAX_SIGNATURE_FONTS
    _sh.rmtree(outside)
with tempfile.TemporaryDirectory() as tmp:     # empty folder -> built-in fallback
    assert docs.signature_fonts(Path(tmp))[0]["file"] == docs.SCRIPT_FONT.name


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

# ------------------------------------------------ BM / RM templates (v1.35)
EXPECT = {"made_on", "region", "partner_name", "passport", "address", "uid", "as_of",
          "sign_partner_name", "sign_passport", "sign_sign_date", "signature"}
for role, title in (("BM", "Business Manager"), ("RM", "Relationship Manager")):
    tpl_path = CM.parent / f"{role}.docx"
    meta = Document(str(tpl_path)).core_properties
    assert meta.author == "Dupoin Markets Ltd" and meta.last_modified_by == "Dupoin Markets Ltd", role
    raw = __import__("zipfile").ZipFile(tpl_path).read("docProps/core.xml").decode()
    assert "ZHUZI" not in raw and "Andrew" not in raw, role               # staff names scrubbed
    want = EXPECT | ({"primary_location"} if role == "BM" else set())    # RM has no location line
    assert set(docs.unfilled_placeholders(tpl_path, signed=True)) >= want, role
    rdata, _ = forms.validate({**AGREEMENT, "role": role})
    o = TMP / f"{role}-signed.docx"
    got = docs.fill_agreement(tpl_path, rdata, o, sign_date="2026-11-02", signature_png=drawn)
    assert set(got) == want | {"allowance", "remark"}, (role, sorted(got))
    assert docs.unfilled_placeholders(o, signed=True) == [], (role, docs.unfilled_placeholders(o, signed=True))
    assert docs.company_column_blank(o), role
    rd = Document(str(o))
    t, pcol = docs._signature_table(rd)
    assert t.rows[0].cells[pcol].text.strip() == f"SIGNED by the {title}", role
    assert t.rows[1].cells[pcol]._tc.xpath(".//*[local-name()='drawing']"), role   # partner column
    rtext = "\n".join(p.text for p in rd.paragraphs)
    made = ("This agreement was made on 1 November 2026" if role == "RM"
            else "This Agreement is made on 1 November 2026")
    assert made in rtext and "made on on" not in rtext, role
    assert "Test Partner" in rtext and "X0000001" in rtext and "1 Sample Road, Sample City" in rtext
    assert "Remark: Rebate will be set by superior" in rtext, role
    assert "Agreed Allowance Scheme:\nNet Deposit : 12%" in rtext, role
    assert "Date\t\t\t: 2 November 2026" in rtext, role
    assert "Dupoin UID|999001" in ["|".join(c.text for c in r.cells) for tb in rd.tables for r in tb.rows]
    assert rd.core_properties.author == "Dupoin Markets Ltd", role
    if role == "BM":
        assert "The primary location is Vietnam, with" in rtext
# the date line regex covers both wordings, and only them
for line, ok in (("This Agreement is made on ", True), ("This agreement was made", True),
                 ("This Agreement is made on", True), ("This Agreement shall be made", False)):
    assert bool(docs.MADE_ON.match(line)) is ok, line

# ------------------------------------- Personal Information Form (admin only)
PIF = docs.pif_template(CM.parent)
assert PIF == CM.parent / "PIF.docx"
raw = __import__("zipfile").ZipFile(PIF).read("docProps/core.xml").decode()
assert "ZHUZI" not in raw and "Andrew" not in raw
assert docs.PIF_UNMAPPED == ["Nickname", "Mobile No", "Attachments", "Date", "Signature"]
po = TMP / "pif.docx"
got = docs.fill_pif(PIF, forms.validate(AGREEMENT)[0], po)
assert got == ["Full Name", "Address", "Passport / National ID Number", "Position",
               "Region / Country", "Email Address", "Dupoin UID", "Passport / ID No"], got
pd = Document(str(po))
cells = {lab: c.text for lab, c in docs._pif_cells(pd)}
assert cells["Full Name"].endswith("Test Partner")
assert cells["Address"].rstrip().endswith("1 Sample Road, Sample City")
assert cells["Passport / National ID Number"].rstrip().endswith("X0000001")
assert cells["Position"].rstrip().endswith("CM") and cells["Region / Country"].rstrip().endswith("Vietnam")
assert cells["Email Address"].rstrip().endswith("partner@example.com")
assert cells["Dupoin UID"].rstrip().endswith("999001")
assert cells["Mobile No"].strip() == "Mobile No:"                 # no form field: blank
assert cells["Nickname"].strip() == "Nickname: -"
assert any(re.search(r"Passport / ID No\s*: X0000001$", p.text.rstrip()) for p in pd.paragraphs)
assert pd.core_properties.author == "Dupoin Markets Ltd"
# missing values stay blank, nothing crashes
got = docs.fill_pif(PIF, {"partner_name": "Only Name"}, TMP / "pif2.docx")
assert got == ["Full Name"], got

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
