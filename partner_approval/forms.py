"""Sales Partner Request form: field definitions, conditional rules, validation
and the confirmation summary. Mirrors the Lark form "Sales PartnerRequest".

The same rules drive the HTML form (which fields are shown for which request
type) and the server-side validation, so the two cannot drift apart.
"""
import datetime as dt
import re

# Request type key -> label. "Dowgrade" is the source form's own spelling.
REQUEST_TYPES = {
    "agreement": "Agreement",
    "special_mib": "Special MIB",
    "downgrade": "Dowgrade",
    "exit": "Remove - Exit",
}
AGREEMENT_TYPES = ("agreement", "special_mib")      # an agreement may be generated
CHANGE_TYPES = ("downgrade", "exit")                 # approve / reject only

ROLES = ("CM", "BM", "RM", "MIB", "NCIB", "RD")
REGIONS = ("Vietnam", "Thailand", "India", "Japan", "Egypt", "PH", "Laos", "Korea",
           "Singapore", "Pakistan", "Bangladesh", "UAE", "Saudi Arabia")
DESKS = ("TH & LA Desk", "VN Desk", "ASIA Desk", "GEM Desk")
# value -> label shown in the form (Lark shows the coloured dots)
SALES_TEAMS = {"Frank Team": "\U0001F7E2 Frank Team",
               "Kevvin Team": "\U0001F7E3 Kevvin Team",
               "Shawn Team": "\U0001F7E1 Shawn Team"}
CONTRACT_TYPES = {"NEW": "NEW", "RENEW / UPDATE": "RENEW / UPDATE"}
NEW_POSITIONS = ("RM", "MIB", "Client")
DOWNGRADE_TYPES = ("Remove from Salary List", "Downgrade CRM Role", "Adjustment Rebate")
YES_NO = ("Yes", "No")

A, C = set(AGREEMENT_TYPES), set(CHANGE_TYPES)
ALL = A | C

SHORT_MAX = 200
LONG_MAX = 5000


def _opts(values):
    if isinstance(values, dict):
        return list(values.items())
    return [(v, v) for v in values]


# Order = the order shown in the form. `req` = request types where the field is
# required, `opt` = types where it is shown but optional. Hidden otherwise.
FIELDS = [
    dict(name="contract_type", label="Contract Type", kind="radio",
         options=_opts(CONTRACT_TYPES), req=A, opt=set()),
    dict(name="sales_team", label="Which Sales Team", kind="radio",
         options=_opts(SALES_TEAMS), req=A, opt=set()),
    dict(name="sales_uid", label="Sales UID", kind="int", req=A, opt=set()),
    dict(name="role", label="Role", kind="radio", options=_opts(ROLES), req=ALL, opt=set()),
    dict(name="uid", label="UID", kind="int", req=ALL, opt=set()),
    dict(name="partner_name", label="Partner Name", kind="text", req=ALL, opt=set()),
    dict(name="passport", label="Passport/National ID", kind="text", req=A, opt=set()),
    dict(name="address", label="Address", kind="textarea", req=A, opt=set()),
    dict(name="team_upline", label="Team Upline ONLY", kind="text", req=A, opt=set(),
         help="111111-222222-333333",
         placeholder="ISS L\u00ca TH\u00c0NH T\u00c2M 446157 > RM 543936 LODIS"),
    dict(name="background", label="Background", kind="textarea", req=A, opt=set()),
    dict(name="allowance_scheme", label="Allowance Scheme", kind="textarea", req=A, opt=set(),
         placeholder="Net Deposit : 12%"),
    dict(name="remarks", label="Remarks", kind="textarea", req=set(), opt=A,
         placeholder="Rebate will be set by superior"),
    dict(name="region", label="Region", kind="select", options=_opts(REGIONS), req=A, opt=set()),
    dict(name="effective_date", label="Effective Date", kind="date", req=ALL, opt=set(),
         help="Apply for Agreement , Special MIB & Downgrade", default_today=True),
    dict(name="onboard_date", label="Onboard Date", kind="date", req=A, opt=set()),
    dict(name="submission_date", label="Agreement Submission Date", kind="date", req=A, opt=set()),
    dict(name="new_position", label="New Position", kind="radio", options=_opts(NEW_POSITIONS),
         req=C, opt=set()),
    dict(name="downgrade_type", label="Type of Downgrade", kind="radio",
         options=_opts(DOWNGRADE_TYPES), req=C, opt=set()),
    # Required for agreements: the signing link is e-mailed there.
    dict(name="partner_email", label="Partner Email", kind="email", req=A, opt=C),
    dict(name="desk", label="Desk Accountable", kind="radio", options=_opts(DESKS),
         req=ALL, opt=set()),
    dict(name="create_agreement", label="Do u want to create agreement?", kind="radio",
         options=_opts(YES_NO), req=set(), opt=ALL),
]
FIELD = {f["name"]: f for f in FIELDS}

_EMAIL = re.compile(r"^[^@\s<>\"']{1,64}@[A-Za-z0-9.-]{1,190}\.[A-Za-z]{2,}$")
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def fields_for(rtype):
    return [f for f in FIELDS if rtype in f["req"] or rtype in f["opt"]]


def is_required(field, rtype):
    return rtype in field["req"]


def _clean(value, multiline):
    value = _CTRL.sub("", (value or "").replace("\r\n", "\n").replace("\r", "\n"))
    if not multiline:
        value = " ".join(value.split())
    return value.strip()


def validate(form):
    """Return (data, errors). `form` is any mapping (request.form).

    Only the fields that belong to the chosen request type are kept; anything
    else that was posted is dropped, so hidden fields can never sneak in.
    """
    errors = {}
    rtype = (form.get("request_type") or "").strip()
    if rtype not in REQUEST_TYPES:
        return {}, {"request_type": "Choose what request you have."}
    data = {"request_type": rtype}
    for f in fields_for(rtype):
        name, kind = f["name"], f["kind"]
        raw = _clean(form.get(name), kind == "textarea")
        if not raw:
            if is_required(f, rtype):
                errors[name] = f"{f['label']} is required."
            continue
        limit = LONG_MAX if kind == "textarea" else SHORT_MAX
        if len(raw) > limit:
            errors[name] = f"{f['label']} is too long (max {limit} characters)."
            continue
        if kind in ("radio", "select"):
            if raw not in {v for v, _ in f["options"]}:
                errors[name] = f"Choose a valid {f['label']}."
                continue
        elif kind == "int":
            if not re.fullmatch(r"\d{1,15}", raw) or int(raw) == 0:
                errors[name] = f"{f['label']} must be a whole number."
                continue
            raw = str(int(raw))
        elif kind == "date":
            try:
                raw = dt.date.fromisoformat(raw).isoformat()
            except ValueError:
                errors[name] = f"{f['label']} must be a date (YYYY-MM-DD)."
                continue
        elif kind == "email":
            if not _EMAIL.match(raw):
                errors[name] = f"{f['label']} is not a valid e-mail address."
                continue
        data[name] = raw
    return data, errors


# ------------------------------------------------------------------ summary
def us_date(iso):
    """2026-11-01 -> 11/01/2026 (the format of the Lark summary)."""
    if not iso:
        return ""
    d = dt.date.fromisoformat(iso)
    return f"{d.month:02d}/{d.day:02d}/{d.year}"


AGREEMENT_TYPE_LABEL = {"NEW": "New Join", "RENEW / UPDATE": "Renew / Update"}


def _block(lines, label, value):
    lines.append(f"{label}:")
    lines.extend((value or "-").split("\n"))


def summary(data):
    """Plain-text confirmation summary, the same text the popup and pages show."""
    rtype = data["request_type"]
    g = data.get
    lines = []
    if rtype in AGREEMENT_TYPES:
        lines.append("Partnership Agreement" if rtype == "agreement"
                     else "Partnership Agreement (Special MIB)")
        lines.append(f"Effective Date: {us_date(g('effective_date'))}")
        lines.append(f"Onboard Date: {us_date(g('onboard_date'))}")
        lines.append(f"Team: {(g('sales_team') or '').replace(' Team', '')}")
        lines.append(f"Region: {g('region', '')}")
        _block(lines, "Team Upline", g("team_upline"))
        _block(lines, "Background", g("background"))
        lines.append("Agreement Type: "
                     f"{AGREEMENT_TYPE_LABEL.get(g('contract_type'), g('contract_type', ''))}")
        _block(lines, "Allowance Scheme", g("allowance_scheme"))
        _block(lines, "Remarks", g("remarks"))
        lines += ["", "Partner details",
                  f"Role: {g('role', '')}",
                  f"UID: {g('uid', '')}",
                  f"Partner Name: {g('partner_name', '')}",
                  f"Passport/National ID: {g('passport', '')}"]
        _block(lines, "Address", g("address"))
        lines += [f"Partner Email: {g('partner_email', '')}",
                  f"Sales UID: {g('sales_uid', '')}",
                  f"Agreement Submission Date: {us_date(g('submission_date'))}",
                  f"Desk Accountable: {g('desk', '')}",
                  f"Create agreement: {g('create_agreement') or '-'}"]
    else:
        lines.append(REQUEST_TYPES[rtype])
        lines += [f"Effective Date: {us_date(g('effective_date'))}",
                  f"Role: {g('role', '')}",
                  f"UID: {g('uid', '')}",
                  f"Partner Name: {g('partner_name', '')}",
                  f"New Position: {g('new_position', '')}",
                  f"Type of Downgrade: {g('downgrade_type', '')}",
                  f"Partner Email: {g('partner_email') or '-'}",
                  f"Desk Accountable: {g('desk', '')}",
                  f"Create agreement: {g('create_agreement') or '-'}"]
    return "\n".join(lines)
