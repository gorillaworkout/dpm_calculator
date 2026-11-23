"""CSV and Excel export of the responses of one event."""
import csv
import datetime as dt
import io

from openpyxl import Workbook
from openpyxl.styles import Font

from .rules import WIB, format_wib

HEADERS = ["Full name", "Email", "Attending", "Food allergy", "Allergy notes", "Submissions",
           "First submitted (WIB)", "Last updated (WIB)"]
WIDTHS = [28, 32, 11, 13, 40, 12, 24, 24]


def _guard(text):
    """Spreadsheet formula-injection guard for text cells."""
    text = text or ""
    return "'" + text if text[:1] in ("=", "+", "-", "@", "\t", "\r") else text


def _cells(r):
    return [_guard(r["full_name"]), _guard(r["email"]), "Yes" if r["attending"] else "No",
            "" if r["has_allergy"] is None else ("Yes" if r["has_allergy"] else "No"),
            _guard(r["allergy_notes"]), r["submission_count"],
            format_wib(r["created_at"]), format_wib(r["updated_at"])]


def to_csv(rows):
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow(HEADERS)
    for r in rows:
        w.writerow(_cells(r))
    return ("\ufeff" + buf.getvalue()).encode("utf-8")


def to_xlsx(rows, sheet_name):
    wb = Workbook()
    ws = wb.active
    bad = '[]:*?/\\'
    ws.title = ("".join(ch for ch in sheet_name if ch not in bad).strip() or "Responses")[:31]
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for r in rows:
        ws.append(_cells(r))
    for i, width in enumerate(WIDTHS):
        ws.column_dimensions[chr(65 + i)].width = width
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def filename(slug, ext, now=None):
    day = dt.datetime.fromtimestamp(now, WIB) if now else dt.datetime.now(WIB)
    return f"dpm-invitation-{slug}-{day:%Y%m%d}.{ext}"
