"""Agreement documents: fill the Word template, convert to PDF, signature images.

The template is filled by finding its own label text ("Name\t: ", "The primary
location is\t\t\t,", the "SIGNED by the ..." table, ...), not by paragraph
number, so a new template for another role works as long as it keeps the same
labels. Formatting is preserved: values are written into new runs that copy the
formatting of the label run next to them.

The company signature column ('SIGNED by the Company', 'Authorized Signature
Dupoin Markets Ltd ____') is never touched: it stays blank.
"""
import base64
import copy
import datetime as dt
import io
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
from pathlib import Path

HERE = Path(__file__).resolve().parent
SIGNATURE_FONTS = HERE / "static" / "signature_fonts"
SCRIPT_FONT = SIGNATURE_FONTS / "1-DarceyOliver-Regular.ttf"  # fallback if a custom folder is empty
MAX_SIGNATURE_FONTS = 8
_FONT_FILE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,63}\.(ttf|otf)$")
MAX_SIGNATURE_BYTES = 2 * 1024 * 1024


class ConversionUnavailable(RuntimeError):
    """LibreOffice is not installed (or not on PATH)."""


class ConversionFailed(RuntimeError):
    pass


class SignatureError(ValueError):
    pass


# ------------------------------------------------------------------ templates
# Roles that have no Word file yet borrow the closest existing one (Bayu, 4 Nov
# 2026: "template yang belum ada pakai template yang ada dulu aja"). A real file
# always wins: drop MIB.docx / NCIB.docx / RD.docx into the templates folder and
# it is used from the next approval on, with no code change. To change a choice,
# edit this dict; to stop borrowing for a role, remove its entry.
TEMPLATE_FALLBACK = {
    "MIB": "RM",    # introducing broker: individual sales role, client and referral
    "NCIB": "RM",   # rebates, no team target, no reporting line -> Relationship Manager
    "RD": "CM",     # senior management with a team target -> Country Manager
}

# What each template calls its party, and what each role is called in a borrowed
# template. "Master IB" is how CM.docx itself names MIBs; NCIB has no confirmed
# long title, so the abbreviation is the defined term ("referred to as "NCIB"").
TEMPLATE_TITLES = {"CM": "Country Manager", "BM": "Business Manager",
                   "RM": "Relationship Manager"}
ROLE_TITLES = {"MIB": "Master IB", "NCIB": "NCIB", "RD": "Regional Director"}


def _real_template(templates_dir, role):
    for name in (f"{role}.docx", f"{role.upper()}.docx", f"{role.lower()}.docx"):
        p = Path(templates_dir) / name
        if p.is_file():
            return p
    return None


def template_info(templates_dir, role):
    """(path, borrowed_from) for a role: borrowed_from is None for the role's own
    file, or the role whose file is used for now (e.g. "RM"). (None, None) when
    there is nothing to use."""
    if not role or not re.fullmatch(r"[A-Za-z]{1,10}", role) or role.upper() == "PIF":
        return None, None                # PIF.docx is the admin-only form, not an agreement
    own = _real_template(templates_dir, role)
    if own is not None:
        return own, None
    src = TEMPLATE_FALLBACK.get(role.upper())
    borrowed = _real_template(templates_dir, src) if src else None
    return (borrowed, src) if borrowed is not None else (None, None)


def template_for(templates_dir, role):
    """The .docx for a role (CM.docx, BM.docx, ...): its own file, else the
    TEMPLATE_FALLBACK one, else None."""
    return template_info(templates_dir, role)[0]


def role_swaps(template, role):
    """[(old, new), ...] that turn a borrowed template's role wording into the
    request's role. Empty when the template is the role's own file."""
    src = Path(template).stem.upper()
    role = (role or "").upper()
    if not role or src == role or src not in TEMPLATE_TITLES or role not in ROLE_TITLES:
        return []
    title, abbr = ROLE_TITLES[role], role
    an = "an" if abbr[0] in "AEFHILMNORSX" else "a"   # spoken letters: "an MIB", "an NCIB", "a RD"
    swaps = []
    if src == "CM" and role == "RD":
        # CM says "report directly to the Regional Director, or in their absence,
        # to the Director of Business Development": an RD cannot report to itself.
        swaps.append(("report directly to the Regional Director, or in their absence, to the "
                      "Director of Business Development",
                      "report directly to the Director of Business Development"))
    swaps += [(TEMPLATE_TITLES[src], title),
              (f"a {src}'s", f"{an} {abbr}'s"),
              (f"the {src} ", f"the {abbr} "), (f"the {src}.", f"the {abbr}.")]
    return swaps


def _replace_across_runs(paragraph, old, new):
    """Replace `old` in a paragraph even when Word split it over several runs.
    The replacement takes the formatting of the run where the match starts."""
    changed = 0
    while True:
        runs = paragraph.runs
        text = "".join(r.text for r in runs)
        i = text.find(old)
        if i < 0 or not runs:
            return changed
        j, pos = i + len(old), 0
        first = True
        for r in runs:
            a, b = pos, pos + len(r.text)
            pos = b
            if b <= i or a >= j:
                continue
            lo, hi = max(a, i) - a, min(b, j) - a
            r.text = r.text[:lo] + (new if first else "") + r.text[hi:]
            first = False
        changed += 1


def _all_paragraphs(doc):
    def walk(container):
        for p in container.paragraphs:
            yield p
        for t in container.tables:
            for row in t.rows:
                for c in row.cells:
                    yield from walk(c)
    yield from walk(doc)
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            yield from walk(part)


def swap_role_wording(doc, swaps):
    """Apply role_swaps() to every paragraph (body, tables, headers, footers)."""
    n = 0
    if swaps:
        for p in _all_paragraphs(doc):
            for old, new in swaps:
                if old in p.text:
                    n += _replace_across_runs(p, old, new)
    return n


def long_date(iso_or_date):
    if not iso_or_date:
        return ""
    d = iso_or_date if isinstance(iso_or_date, dt.date) else dt.date.fromisoformat(iso_or_date)
    return f"{d.day} {d.strftime('%B')} {d.year}"


# ---------------------------------------------------------------- docx helpers
def _copy_run_after(paragraph, model, text):
    """New run with `model`'s formatting appended to `paragraph`."""
    run = paragraph.add_run("")
    if model is not None and model._r.rPr is not None:
        run._r.insert(0, copy.deepcopy(model._r.rPr))
    _set_text(run, text)
    return run


def _set_text(run, text):
    """Run text with line breaks for '\n' (a plain .text would show a box)."""
    lines = (text or "").split("\n")
    run.text = lines[0]
    for line in lines[1:]:
        run.add_break()
        run._r.append(_t(line))


def _t(text):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    return t


def _last_text_run(paragraph):
    runs = [r for r in paragraph.runs if r.text.strip()]
    return runs[-1] if runs else (paragraph.runs[-1] if paragraph.runs else None)


def _append_value(paragraph, value):
    if value:
        _copy_run_after(paragraph, _last_text_run(paragraph), value)


def _fill_tab_gap(paragraph, anchor, value):
    """'The primary location is\t\t\t,' -> 'The primary location is Vietnam,'.

    The tab-only runs right after the run that ends with `anchor` are replaced
    by the value. Returns True when the gap was found.
    """
    runs = paragraph.runs
    for i, run in enumerate(runs):
        if run.text.rstrip().endswith(anchor):
            j = i + 1
            gap = []
            while j < len(runs) and runs[j].text and not runs[j].text.strip("\t"):
                gap.append(runs[j])
                j += 1
            if gap:
                gap[0].text = " " + value if not run.text.endswith(" ") else value
                for g in gap[1:]:
                    g.text = ""
                return True
    return False


def _insert_paragraph_after(paragraph, text, model_run=None):
    from docx.text.paragraph import Paragraph
    new_p = copy.deepcopy(paragraph._p)
    for child in list(new_p):
        if child.tag != new_p.tag.replace("}p", "}pPr"):
            new_p.remove(child)
    paragraph._p.addnext(new_p)
    para = Paragraph(new_p, paragraph._parent)
    _copy_run_after(para, model_run, text)
    return para


def _iter_paragraphs(doc):
    """Body paragraphs in order, with the index of the table that precedes them."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def _signature_table(doc):
    for t in doc.tables:
        heads = [c.text.strip() for c in t.rows[0].cells] if t.rows else []
        if any(h.startswith("SIGNED by the Company") for h in heads):
            partner_col = next(i for i, h in enumerate(heads)
                               if h.startswith("SIGNED by the") and "Company" not in h)
            return t, partner_col
    return None, None


# The opening date line. Templates differ: "This Agreement is made on " (CM, BM)
# and "This agreement was made" (RM). Group 1 = the template already has "on".
MADE_ON = re.compile(r"^This [Aa]greement (?:is|was) made( on)?\s*$")
LABEL = r"\s*:\s*$"
PARTY_LABELS = {  # label regex (whole paragraph) -> data key
    r"^Region/Country" + LABEL: "region",
    r"^Name" + LABEL: "partner_name",
    r"^Passport/National ID" + LABEL: "passport",
    r"^Address" + LABEL: "address",
}
SIGN_LABELS = {  # signature block lines under the SIGNED table (end of the line)
    r"Name" + LABEL: "partner_name",
    r"Passport/ID No" + LABEL: "passport",
    r"Date" + LABEL: "sign_date",
}


def fill_agreement(template, data, out_path, sign_date=None, signature_png=None):
    """Write the filled agreement to `out_path` (.docx). Returns a list of the
    placeholders that were filled (for tests and the audit trail)."""
    from docx import Document
    from docx.shared import Inches

    doc = Document(str(template))
    swaps = role_swaps(template, data.get("role"))
    swapped = swap_role_wording(doc, swaps)
    values = {
        "effective_date": long_date(data.get("effective_date")),
        "region": data.get("region", ""),
        "partner_name": data.get("partner_name", ""),
        "passport": data.get("passport", ""),
        "address": ", ".join(x.strip() for x in (data.get("address") or "").split("\n") if x.strip()),
        "uid": data.get("uid", ""),
        "sign_date": long_date(sign_date) if sign_date else "",
    }
    filled = []
    sig_table, partner_col = _signature_table(doc)
    after_sig_table = False
    party_done = set()

    for block in list(_iter_paragraphs(doc)):
        if not hasattr(block, "runs"):                 # a table
            if sig_table is not None and block._tbl is sig_table._tbl:
                after_sig_table = True
            continue
        p = block
        text = p.text
        m = MADE_ON.match(text)
        if m and "made_on" not in filled:
            # CM/BM: "This Agreement is made on " -> "... made on 1 November 2026"
            # RM:    "This agreement was made"    -> "... was made on 1 November 2026"
            lead = "" if text.endswith((" ", "\t")) else " "
            date = values["effective_date"]
            _append_value(p, (lead + date) if m.group(1) else f"{lead}on {date}")
            filled.append("made_on")
            continue
        if not after_sig_table:
            for rx, key in PARTY_LABELS.items():
                if key not in party_done and re.match(rx, text):
                    _append_value(p, values[key])
                    party_done.add(key)
                    filled.append(key)
                    break
        if "primary location is" in text and _fill_tab_gap(p, "primary location is", values["region"]):
            filled.append("primary_location")
        if re.search(r"as of\t+,", text) and _fill_tab_gap(p, "as of", values["effective_date"]):
            filled.append("as_of")
        if after_sig_table:
            for rx, key in SIGN_LABELS.items():
                if re.search(rx, text) and f"sign_{key}" not in filled:
                    _append_value(p, values[key])
                    filled.append(f"sign_{key}")
                    break
        if re.match(r"^Allowance Scheme\b", text) and "allowance" not in filled:
            scheme = (data.get("allowance_scheme") or "").strip()
            if scheme:
                _insert_paragraph_after(p, "\nAgreed Allowance Scheme:\n" + scheme,
                                        _last_text_run(p))
            filled.append("allowance")
        if re.match(r"^Remark:", text) and "remark" not in filled:
            remarks = (data.get("remarks") or "").strip()
            if remarks:
                first = p.runs[0]
                run = _copy_run_after(p, first, remarks)
                run.bold = False                      # label bold, value plain
                first._r.addnext(run._r)              # right after "Remark: "
            filled.append("remark")

    for t in doc.tables:
        for row in t.rows:
            cells = row.cells
            for i, c in enumerate(cells[:-1]):
                if c.text.strip() == "Dupoin UID":
                    target = cells[i + 1]
                    para = target.paragraphs[0]
                    if para.runs:
                        para.runs[0].text = values["uid"]
                        for r in para.runs[1:]:
                            r.text = ""
                    else:
                        para.add_run(values["uid"])
                    filled.append("uid")

    if signature_png is not None and sig_table is not None:
        cell = sig_table.rows[1].cells[partner_col]
        para = cell.paragraphs[0]
        para.add_run().add_picture(io.BytesIO(signature_png), width=Inches(1.9))
        filled.append("signature")

    if swapped:
        filled.append("role_wording")
    cp = doc.core_properties
    cp.author = "Dupoin Markets Ltd"
    cp.last_modified_by = "DW Calculator"
    cp.comments = ""
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return filled


def unfilled_placeholders(docx_path, signed=False):
    """Labels that still have no value. Empty list = everything filled."""
    from docx import Document
    doc = Document(str(docx_path))
    left = []
    sig_table, partner_col = _signature_table(doc)
    after = False
    for block in _iter_paragraphs(doc):
        if not hasattr(block, "runs"):
            if sig_table is not None and block._tbl is sig_table._tbl:
                after = True
            continue
        text = block.text
        if MADE_ON.match(text):
            left.append("made_on")
        if not after:
            for rx, key in PARTY_LABELS.items():
                if re.match(rx, text):
                    left.append(key)
        if re.search(r"primary location is\t+,", text):
            left.append("primary_location")
        if re.search(r"as of\t+,", text):
            left.append("as_of")
        if after:
            for rx, key in SIGN_LABELS.items():
                if re.search(rx, text) and (key != "sign_date" or signed):
                    left.append(f"sign_{key}")
    for t in doc.tables:
        for row in t.rows:
            cells = row.cells
            for i, c in enumerate(cells[:-1]):
                if c.text.strip() == "Dupoin UID" and cells[i + 1].text.strip() in ("", "111111"):
                    left.append("uid")
    if signed and sig_table is not None:
        cell = sig_table.rows[1].cells[partner_col]
        if not cell._tc.xpath(".//*[local-name()='drawing']"):
            left.append("signature")
    return left


def company_column_blank(docx_path):
    """True when the 'SIGNED by the Company' cell has no text and no picture."""
    from docx import Document
    doc = Document(str(docx_path))
    t, partner_col = _signature_table(doc)
    company = 1 - partner_col
    cell = t.rows[1].cells[company]
    return not cell.text.strip() and not cell._tc.xpath(".//*[local-name()='drawing']")


# ----------------------------------------------------------------- PDF convert
def find_soffice(configured=""):
    """Path of LibreOffice, or None. DW_SOFFICE, when set, is the only candidate."""
    if configured:
        found = shutil.which(configured)
        if found:
            return found
        return configured if os.path.isfile(configured) and os.access(configured, os.X_OK) else None
    return shutil.which("soffice") or shutil.which("libreoffice")


SOFFICE_MISSING = ("PDF conversion is not available: LibreOffice ('soffice') was not found "
                   "on this server. Install it with 'sudo apt install libreoffice-writer' "
                   "(or set DW_SOFFICE to its path), then click 'Retry agreement'.")


def docx_to_pdf(docx_path, pdf_path, soffice="", timeout=180):
    """Convert one .docx to .pdf. The ONLY place that knows how; swap it here.

    Raises ConversionUnavailable when LibreOffice is missing and ConversionFailed
    when it ran but produced no PDF.
    """
    exe = find_soffice(soffice)
    if not exe:
        raise ConversionUnavailable(SOFFICE_MISSING)
    docx_path, pdf_path = Path(docx_path), Path(pdf_path)
    with tempfile.TemporaryDirectory(prefix="dw-soffice-") as tmp:
        # Own profile per call: two gunicorn workers converting at once would
        # otherwise fight over the same LibreOffice profile lock.
        profile = Path(tmp, "profile").as_uri()
        try:
            proc = subprocess.run(
                [exe, f"-env:UserInstallation={profile}", "--headless", "--norestore",
                 "--convert-to", "pdf", "--outdir", tmp, str(docx_path)],
                capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            raise ConversionFailed(f"LibreOffice did not finish within {timeout} s.")
        produced = Path(tmp, docx_path.stem + ".pdf")
        if not produced.is_file() or produced.stat().st_size == 0:
            raise ConversionFailed("LibreOffice did not produce a PDF: "
                                   + (proc.stderr or proc.stdout or "").strip()[-400:])
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), pdf_path)
    return pdf_path


# ------------------------------------------------------------------ signatures
PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def drawn_signature(data_url):
    """Canvas data URL -> cleaned, cropped PNG bytes. Raises SignatureError."""
    from PIL import Image
    if not data_url or not data_url.startswith("data:image/png;base64,"):
        raise SignatureError("Please draw your signature first.")
    raw = data_url.split(",", 1)[1]
    if len(raw) > MAX_SIGNATURE_BYTES * 4 // 3 + 16:
        raise SignatureError("The signature image is too large.")
    try:
        blob = base64.b64decode(raw, validate=True)
    except ValueError:
        raise SignatureError("The signature could not be read. Please draw it again.")
    if not blob.startswith(PNG_MAGIC):
        raise SignatureError("The signature could not be read. Please draw it again.")
    try:
        img = Image.open(io.BytesIO(blob))
        if img.width * img.height > 4000 * 2000:
            raise SignatureError("The signature image is too large.")
        img = img.convert("RGBA")
    except SignatureError:
        raise
    except Exception:
        raise SignatureError("The signature could not be read. Please draw it again.")
    return _crop(img, "Please draw your signature first.")


def _font_label(stem):
    stem = re.sub(r"^\d+[-_]", "", stem)                         # "1-Foo" -> "Foo" (sort prefix)
    stem = re.sub(r"[-_](Regular|VariableFont.*|Variable)$", "", stem, flags=re.I)
    stem = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", stem)              # CamelCase -> words
    return " ".join(stem.replace("_", " ").replace("-", " ").split()) or "Script"


def signature_fonts(folder=None):
    """Fonts offered on the 'Type name' tab: [{id, label, file}], first = default.

    Only plain file names (letters, digits, - and _) ending in .ttf/.otf directly
    inside the folder count; symlinks and anything else are ignored."""
    folder = Path(folder or SIGNATURE_FONTS)
    out = []
    if folder.is_dir():
        for f in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
            if _FONT_FILE.match(f.name) and f.is_file() and not f.is_symlink():
                out.append({"id": f.stem, "label": _font_label(f.stem), "file": f.name,
                            "path": f})
    if not out and SCRIPT_FONT.is_file():
        out.append({"id": SCRIPT_FONT.stem, "label": _font_label(SCRIPT_FONT.stem),
                    "file": SCRIPT_FONT.name, "path": SCRIPT_FONT})
    out = out[:MAX_SIGNATURE_FONTS]
    for f in out:
        f["chars"] = "".join(sorted(font_coverage(f.pop("path"))))
    return out


# Characters a typed name may contain. Each font's real coverage of this set is
# measured once (font_coverage); the page disables a style that cannot show the
# typed name and the server refuses it with the same rule -> never a tofu box.
SIGNATURE_CHARSET = frozenset(
    [chr(c) for c in range(0x21, 0x7F)] + [chr(c) for c in range(0xC0, 0x250)]
    + [chr(c) for c in range(0x1E00, 0x1F00)] + list("\u2018\u2019\u201c\u201d"))
_COVERAGE = {}


def font_coverage(path):
    """Characters of SIGNATURE_CHARSET that the font really draws (not the tofu box)."""
    from PIL import ImageFont
    path = Path(path)
    st = path.stat()
    key = (str(path.resolve()), st.st_mtime_ns, st.st_size)
    if key not in _COVERAGE:
        font = ImageFont.truetype(str(path), 48)
        missing = set(_missing_glyphs(font, "".join(SIGNATURE_CHARSET)))
        _COVERAGE[key] = frozenset(SIGNATURE_CHARSET - missing)
    return _COVERAGE[key]


def unsupported_chars(name, font_id=None, folder=None):
    """Letters of `name` the chosen style cannot show, in order of appearance."""
    name = unicodedata.normalize("NFC", name or "")
    have = font_coverage(signature_font_path(font_id, folder))
    out = []
    for ch in name:
        if not ch.isspace() and ch not in have and ch not in out:
            out.append(ch)
    return out


def first_font_for(name, folder=None):
    """Id of the first offered style that can show `name` (None if none can)."""
    for f in signature_fonts(folder):
        if not unsupported_chars(name, f["id"], folder):
            return f["id"]
    return None


def signature_font_path(font_id=None, folder=None):
    """Path of an offered font. Unknown / empty id -> the default font.
    The id is only ever matched against the folder listing, never joined as a path."""
    folder = Path(folder or SIGNATURE_FONTS)
    fonts = signature_fonts(folder)
    if not fonts:
        raise SignatureError("Typed signatures are not available. Please draw your signature.")
    pick = next((f for f in fonts if f["id"] == font_id), None) if font_id else fonts[0]
    if pick is None:
        raise SignatureError("Choose one of the signature styles.")
    base = folder if (folder / pick["file"]).is_file() else SCRIPT_FONT.parent
    path = (base / pick["file"]).resolve()
    if path.parent != base.resolve():
        raise SignatureError("Choose one of the signature styles.")
    return path


def typed_signature(name, font_id=None, folder=None):
    """Typed name rendered in the chosen script font -> PNG bytes."""
    from PIL import Image, ImageDraw, ImageFont
    name = " ".join(unicodedata.normalize("NFC", name or "").split())
    if len(name) < 2:
        raise SignatureError("Please type your full name.")
    if len(name) > 80:
        raise SignatureError("The name is too long.")
    path = signature_font_path(font_id, folder)
    if unsupported_chars(name, font_id, folder):
        raise SignatureError("This signature style cannot show some letters of your name. "
                             "Please choose another style or draw your signature.")
    font = ImageFont.truetype(str(path), 96)
    left, top, right, bottom = font.getbbox(name)
    img = Image.new("RGBA", (right - left + 80, bottom - top + 80), (255, 255, 255, 0))
    ImageDraw.Draw(img).text((40 - left, 40 - top), name, font=font, fill=(16, 32, 96, 255))
    return _crop(img, "Please type your full name.")


def _missing_glyphs(font, text):
    """Characters the font would draw as the empty 'tofu' box."""
    def mask(ch):
        m = font.getmask(ch)
        return m.size, bytes(m)
    tofu = {mask("\uffff"), mask("\U0010fffd")}
    return [ch for ch in set(text) if not ch.isspace() and mask(ch) in tofu]


def _crop(img, empty_message):
    from PIL import Image
    alpha = img.getchannel("A")
    # A canvas may be exported with a white background: treat white as empty too.
    gray = img.convert("L")
    ink = Image.eval(gray, lambda v: 255 if v < 200 else 0)
    mask = Image.composite(ink, Image.new("L", img.size, 0), alpha.point(lambda a: 255 if a > 40 else 0))
    box = mask.getbbox()
    if not box:
        raise SignatureError(empty_message)
    pad = 12
    box = (max(0, box[0] - pad), max(0, box[1] - pad),
           min(img.width, box[2] + pad), min(img.height, box[3] + pad))
    out = Image.new("RGBA", (box[2] - box[0], box[3] - box[1]), (255, 255, 255, 0))
    out.paste(img.crop(box), (0, 0))
    buf = io.BytesIO()
    out.save(buf, "PNG", optimize=True)
    return buf.getvalue()


# --------------------------------------------------------------------- uploads
MAX_UPLOAD = 20 * 1024 * 1024
UPLOAD_TYPES = {"pdf": b"%PDF-", "png": PNG_MAGIC, "jpg": b"\xff\xd8\xff"}


def sniff_upload(head):
    """File type from the first bytes (never from the name). None = refused."""
    for ext, magic in UPLOAD_TYPES.items():
        if head.startswith(magic):
            return ext
    return None


def check_upload(blob):
    """File type of a signed upload, checked by content. None = refused.

    The magic bytes decide the type (never the file name); images must also
    decode and stay below a sane pixel count, PDFs must have a PDF trailer.
    """
    ext = sniff_upload(blob[:16])
    if ext is None:
        return None
    if ext == "pdf":
        return ext if b"%%EOF" in blob[-2048:] else None
    from PIL import Image
    try:
        with Image.open(io.BytesIO(blob)) as img:
            if img.width * img.height > 60_000_000:
                return None
            img.verify()
    except Exception:  # noqa: BLE001 - anything unreadable is refused
        return None
    return ext


# ------------------------------------------------------------- page previews
MAX_PREVIEW_PAGES = 40


def pdf_previews(pdf_path, out_dir, width=1240):
    """Render each PDF page to out_dir/page-<n>.png for the signing page.

    Mobile browsers often cannot show an embedded PDF, page images always work.
    Uses pypdfium2 (installed with pdfplumber). Returns the number of pages
    rendered, 0 when previews are not available -- the page then falls back to
    the embedded PDF and the Open / Download buttons.
    """
    out_dir = Path(out_dir)
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(str(pdf_path))
    except Exception:  # noqa: BLE001 - previews are optional
        return 0
    tmp = Path(tempfile.mkdtemp(prefix="preview-", dir=str(out_dir.parent)))
    try:
        n = min(len(doc), MAX_PREVIEW_PAGES)
        for i in range(n):
            page = doc[i]
            scale = width / max(page.get_width(), 1)
            img = page.render(scale=scale).to_pil().convert("RGB")
            img.save(tmp / f"page-{i + 1}.png", "PNG", optimize=True)
        if out_dir.exists():
            shutil.rmtree(out_dir)
        tmp.rename(out_dir)
        return n
    except Exception:  # noqa: BLE001
        shutil.rmtree(tmp, ignore_errors=True)
        return 0
    finally:
        doc.close()


def preview_count(out_dir):
    out_dir = Path(out_dir)
    n = 0
    while (out_dir / f"page-{n + 1}.png").is_file():
        n += 1
    return n


# ------------------------------------------------- Personal Information Form
# Admin-only (v1.35): made for Agreement requests at approval, stored with the
# request files, never e-mailed, never reachable from a /sign link.
PIF_FILE = "PIF.docx"
# label at the start of a table cell -> request field. None = no form field
# for it yet, left blank for the partner / admin to complete.
PIF_CELLS = (
    ("Full Name", "partner_name"),
    ("Nickname", None),
    ("Address", "address"),
    ("Passport / National ID Number", "passport"),
    ("Position", "role"),
    ("Region / Country", "region"),
    ("Email Address", "partner_email"),
    ("Mobile No", None),
    ("Dupoin UID", "uid"),
    ("Attachments", None),
)
# Declaration lines under the table
PIF_LINES = (("Date", None), ("Signature", None), ("Passport / ID No", "passport"))
PIF_UNMAPPED = [label for label, key in PIF_CELLS + PIF_LINES if key is None]


def pif_template(templates_dir):
    p = Path(templates_dir) / PIF_FILE
    return p if p.is_file() else None


def _pif_cells(doc):
    """[(label, cell)] for the PIF table, merged cells once."""
    if not doc.tables:
        return []
    out, seen = [], set()
    for row in doc.tables[0].rows:
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            head = cell.paragraphs[0].text.strip() if cell.paragraphs else ""
            label = next((lab for lab, _k in PIF_CELLS if head.startswith(lab)), None)
            if label:
                out.append((label, cell))
    return out


def fill_pif(template, data, out_path):
    """Write the filled PIF (.docx). Returns the labels that got a value."""
    from docx import Document
    doc = Document(str(template))
    keys = dict(PIF_CELLS)
    values = {
        "partner_name": data.get("partner_name", ""),
        "address": ", ".join(x.strip() for x in (data.get("address") or "").split("\n") if x.strip()),
        "passport": data.get("passport", ""),
        "role": data.get("role", ""),
        "region": data.get("region", ""),
        "partner_email": data.get("partner_email", ""),
        "uid": data.get("uid", ""),
    }
    filled = []
    for label, cell in _pif_cells(doc):
        key = keys[label]
        value = values.get(key) if key else ""
        if not value or label in filled:
            continue
        head = cell.paragraphs[0]
        model = _last_text_run(head)
        nxt = cell.paragraphs[1] if len(cell.paragraphs) > 1 else None
        if nxt is not None and not nxt.text.strip():
            for r in nxt.runs:
                r.text = ""
            if head._p.pPr is not None:       # same indent as the label line
                if nxt._p.pPr is not None:
                    nxt._p.remove(nxt._p.pPr)
                nxt._p.insert(0, copy.deepcopy(head._p.pPr))
            run = _copy_run_after(nxt, model, value)
        else:                             # one paragraph ending in a line break
            run = _copy_run_after(head, model, value if head.text.endswith("\n") else " " + value)
        run.bold = False
        filled.append(label)
    for p in doc.paragraphs:
        if re.search(r"Passport / ID No\s*$", p.text) and values["passport"] \
                and "Passport / ID No" not in filled:
            run = _copy_run_after(p, _last_text_run(p), ": " + values["passport"])
            run.bold = False
            filled.append("Passport / ID No")
    cp = doc.core_properties
    cp.author = "Dupoin Markets Ltd"
    cp.last_modified_by = "DW Calculator"
    cp.comments = ""
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return filled
