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
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT_FONT = HERE / "static" / "GreatVibes-Regular.ttf"
MAX_SIGNATURE_BYTES = 2 * 1024 * 1024


class ConversionUnavailable(RuntimeError):
    """LibreOffice is not installed (or not on PATH)."""


class ConversionFailed(RuntimeError):
    pass


class SignatureError(ValueError):
    pass


# ------------------------------------------------------------------ templates
def template_for(templates_dir, role):
    """The .docx for a role (CM.docx, BM.docx, ...) or None when not ready yet."""
    if not role or not re.fullmatch(r"[A-Za-z]{1,10}", role):
        return None
    for name in (f"{role}.docx", f"{role.upper()}.docx", f"{role.lower()}.docx"):
        p = Path(templates_dir) / name
        if p.is_file():
            return p
    return None


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
        if re.match(r"^This Agreement is made on\s*$", text):
            _append_value(p, values["effective_date"])
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
        if re.match(r"^This Agreement is made on\s*$", text):
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


def typed_signature(name):
    """Typed name rendered in a script font -> PNG bytes."""
    from PIL import Image, ImageDraw, ImageFont
    name = " ".join((name or "").split())
    if len(name) < 2:
        raise SignatureError("Please type your full name.")
    if len(name) > 80:
        raise SignatureError("The name is too long.")
    font = ImageFont.truetype(str(SCRIPT_FONT), 96)
    left, top, right, bottom = font.getbbox(name)
    img = Image.new("RGBA", (right - left + 80, bottom - top + 80), (255, 255, 255, 0))
    ImageDraw.Draw(img).text((40 - left, 40 - top), name, font=font, fill=(16, 32, 96, 255))
    return _crop(img, "Please type your full name.")


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
