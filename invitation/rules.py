"""Pure validation rules: event fields, slug, WIB closing time, the RSVP form."""
import datetime as dt
import re
import unicodedata

WIB = dt.timezone(dt.timedelta(hours=7), "WIB")
NAME_MAX = 120
EMAIL_MAX = 254
ALLERGY_MAX = 200
RESERVED_SLUGS = {"new", "admin", "api", "image", "thanks", "edit", "static"}
_SLUG = re.compile(r"^[a-z0-9][a-z0-9-]{1,58}[a-z0-9]$")
_EMAIL = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
_CONTROL = re.compile(r"[\x00-\x1f\x7f]")
_SPACES = re.compile(r"[ \t\n\r\f\v\u00a0\u2028\u2029]+")
_WIB_INPUT = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$")


def clean_text(raw):
    """NFKC, collapse ordinary whitespace, trim. Other control characters stay so
    the caller can reject them."""
    return _SPACES.sub(" ", unicodedata.normalize("NFKC", str(raw or ""))).strip()


def slug_issue(slug):
    if not _SLUG.match(slug or "") or "--" in slug:
        return "Use 3–60 lowercase letters, numbers and hyphens (no hyphen at the start or end)."
    if slug in RESERVED_SLUGS:
        return "That word is reserved. Choose another link."
    return None


def suggest_slug(name):
    text = unicodedata.normalize("NFKD", (name or "").replace("đ", "d").replace("Đ", "D"))
    text = "".join(ch for ch in text if not unicodedata.combining(ch)).lower()
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-")[:60].strip("-")


def parse_event_name(raw):
    value = clean_text(raw)
    if len(value) < 2 or len(value) > NAME_MAX or _CONTROL.search(value):
        return None, "Enter a name of 2–120 characters."
    return value, None


def parse_closes_at(raw, now=None):
    """A datetime-local value (YYYY-MM-DDTHH:MM) read as WIB -> unix time.
    Empty clears it (None). Returns (value, error)."""
    value = (raw or "").strip()
    if not value:
        return None, None
    if not _WIB_INPUT.match(value):
        return None, "Enter a valid closing date and time."
    try:
        when = dt.datetime.strptime(value, "%Y-%m-%dT%H:%M").replace(tzinfo=WIB)
    except ValueError:
        return None, "Enter a valid closing date and time."
    ts = when.timestamp()
    if ts <= (now if now is not None else dt.datetime.now(WIB).timestamp()):
        return None, "Closing time must be in the future. To close now, use Close registration."
    return ts, None


def to_wib_input(ts):
    if not ts:
        return ""
    return dt.datetime.fromtimestamp(ts, WIB).strftime("%Y-%m-%dT%H:%M")


def format_wib(ts):
    """'20 Dec 2026, 18:05 WIB', or '' for no time."""
    if not ts:
        return ""
    return dt.datetime.fromtimestamp(ts, WIB).strftime("%d %b %Y, %H:%M WIB")


def is_valid_email(email):
    return len(email) <= EMAIL_MAX and bool(_EMAIL.match(email))


def _yes_no(raw):
    value = str(raw or "").strip().lower()
    return True if value == "yes" else False if value == "no" else None


def parse_response(form):
    """Server-side check of the public form. Returns (value, field_errors).

    Allergy answers are only read when they were asked: attending Yes, then
    allergy Yes. Anything else posted for them is discarded (stored as NULL).
    """
    errors = {}
    full_name = clean_text(form.get("full_name"))
    if len(full_name) < 2 or _CONTROL.search(full_name):
        errors["full_name"] = "Please enter your full name."
    elif len(full_name) > NAME_MAX:
        errors["full_name"] = f"Please use at most {NAME_MAX} characters."

    email = str(form.get("email") or "").strip()
    if not is_valid_email(email):
        errors["email"] = "Please enter a valid email address."

    attending = _yes_no(form.get("attending"))
    if attending is None:
        errors["attending"] = "Please choose Yes or No."

    has_allergy = notes = None
    if attending is True:
        has_allergy = _yes_no(form.get("has_allergy"))
        if has_allergy is None:
            errors["has_allergy"] = "Please choose Yes or No."
        elif has_allergy:
            text = clean_text(form.get("allergy_notes"))
            if not text or _CONTROL.search(text):
                errors["allergy_notes"] = "Please tell us what you are allergic to."
            elif len(text) > ALLERGY_MAX:
                errors["allergy_notes"] = f"Please use at most {ALLERGY_MAX} characters."
            else:
                notes = text
    if errors:
        return None, errors
    return {"full_name": full_name, "email": email, "email_normalized": email.lower(),
            "attending": attending, "has_allergy": has_allergy, "allergy_notes": notes}, {}
