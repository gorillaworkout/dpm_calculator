"""Settings, all from DW_* environment variables, read when the app starts.

DW_APPROVAL_DB          SQLite file (default: approval.db next to users.db)
DW_APPROVAL_DIR         generated agreements, signatures, uploads
                        (default: approval_files/ next to users.db)
DW_APPROVAL_TEMPLATES   folder of <ROLE>.docx agreement templates
                        (default: partner_approval/agreement_templates)
DW_SIGN_LINK_DAYS       signing link lifetime in days (default 14)
DW_PUBLIC_BASE_URL      base of the signing link (default https://dw-calculator.gorillaworkout.id)
DW_SOFFICE              path to LibreOffice soffice (default: found on PATH)
DW_SMTP_HOST / DW_SMTP_PORT / DW_SMTP_USER / DW_SMTP_PASSWORD / DW_SMTP_FROM
DW_SMTP_TLS             starttls (default) | ssl | none
DW_SMTP_FROM_NAME       display name of the sender (default "Dupoin Markets")
DW_SMTP_REPLY_TO        optional Reply-To address (default: none)
"""
import os
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent


def _auth_db():
    return Path(os.environ.get("DW_AUTH_DB", HERE.parent / "users.db"))


@dataclass
class Settings:
    db: Path
    files: Path
    templates: Path
    link_days: float
    base_url: str
    soffice: str
    smtp: dict = field(default_factory=dict)

    @property
    def smtp_ready(self):
        return bool(self.smtp.get("host") and self.smtp.get("from"))


def load():
    env = os.environ.get
    auth_db = _auth_db()
    try:
        days = float(env("DW_SIGN_LINK_DAYS") or 14)
    except ValueError:
        days = 14.0
    if days <= 0:
        days = 14.0
    tls = (env("DW_SMTP_TLS") or "starttls").strip().lower()
    tls = {"1": "starttls", "true": "starttls", "yes": "starttls", "0": "none",
           "false": "none", "no": "none", "off": "none"}.get(tls, tls)
    port = env("DW_SMTP_PORT") or ("465" if tls == "ssl" else "587")
    return Settings(
        db=Path(env("DW_APPROVAL_DB") or auth_db.with_name("approval.db")).expanduser(),
        files=Path(env("DW_APPROVAL_DIR") or auth_db.with_name("approval_files")).expanduser(),
        templates=Path(env("DW_APPROVAL_TEMPLATES") or HERE / "agreement_templates").expanduser(),
        link_days=days,
        base_url=(env("DW_PUBLIC_BASE_URL") or "https://dw-calculator.gorillaworkout.id").rstrip("/"),
        soffice=env("DW_SOFFICE") or "",
        smtp={"host": env("DW_SMTP_HOST") or "", "port": int(port) if port.isdigit() else 587,
              "user": env("DW_SMTP_USER") or "", "password": env("DW_SMTP_PASSWORD") or "",
              "from": (env("DW_SMTP_FROM") or "").strip(), "tls": tls,
              "from_name": (env("DW_SMTP_FROM_NAME") or "Dupoin Markets").strip(),
              "reply_to": (env("DW_SMTP_REPLY_TO") or "").strip()},
    )
