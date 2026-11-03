"""Signing-link e-mail over SMTP. Nothing is sent when SMTP is not configured.

One branded template (HTML table layout with inline CSS for Gmail, Outlook and
Lark, plus a matching plain-text part) for the first link and every re-sign.
"""
import re
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr, make_msgid, parseaddr
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape

from .forms import valid_email

NOT_CONFIGURED = "Email not sent: SMTP not configured"
SUBJECT = "Dupoin Markets \u2013 Partnership Agreement for Signature"
ADDRESS = "Hamchako, Mutsamudu, Anjouan, Union of Comoros"
COMPANY = "Dupoin Markets Ltd"
LOGO_PATH = "/sign/static/dupoin-logo.png"

_ENV = Environment(loader=FileSystemLoader(str(Path(__file__).parent / "templates" / "approval")),
                   autoescape=select_autoescape(default=True))
_HEADER_BAD = re.compile(r"[\x00-\x1f\x7f\u2028\u2029]")


def header_safe(value, limit=200):
    """One line, no CR/LF or other control characters: no header injection."""
    return " ".join(_HEADER_BAD.sub(" ", value or "").split())[:limit]


def _sender(smtp):
    """From header: DW_SMTP_FROM, with DW_SMTP_FROM_NAME as display name unless
    DW_SMTP_FROM already carries one ("Name <addr>")."""
    raw = header_safe(smtp.get("from"), 320)
    name, addr = parseaddr(raw)
    if not valid_email(addr):
        return None
    if "<" not in raw:
        name = smtp.get("from_name") or "Dupoin Markets"
    return formataddr((header_safe(name, 80), addr))


def content(settings, partner_name, link, expires_label, resign_note=None):
    """(subject, text, html) -- also used by the screenshot/preview tooling."""
    name = header_safe(partner_name, 120) or "Partner"
    reply_to = settings.smtp.get("reply_to") or ""
    reply_to = reply_to if valid_email(reply_to) else ""
    if resign_note:
        heading = "An updated agreement is ready for your signature"
        intro = (f"We have prepared an updated version of your partnership agreement with "
                 f"{COMPANY}. Please sign the updated agreement using the new link below; "
                 "earlier links no longer work.")
    else:
        heading = "Your agreement is ready for signature"
        intro = (f"Thank you for partnering with {COMPANY}. Your partnership agreement has been "
                 "approved and is ready for your signature.")
    html = _ENV.get_template("email_signing.html").render(
        subject=SUBJECT, preheader=f"Please review and sign by {expires_label}.",
        heading=heading, intro=intro, name=name, note=resign_note, link=link,
        expires=expires_label, address=ADDRESS, reply_to=reply_to,
        logo_url=f"{settings.base_url}{LOGO_PATH}")
    note_txt = f"\nNote from Dupoin: {resign_note}\n" if resign_note else ""
    text = (
        f"{heading}\n\n"
        f"Dear {name},\n\n"
        f"{intro}\n{note_txt}\n"
        "Please review the agreement and sign it online. You can sign on screen, or download "
        "the PDF, sign it by hand and upload the signed copy.\n\n"
        f"Review & Sign Agreement:\n{link}\n\n"
        f"This link is personal and can be used once. It expires on {expires_label}.\n\n"
        "Kind regards,\nDupoin Markets Partnerships Team\n\n"
        "--\n"
        f"{COMPANY}\n{ADDRESS}\n\n"
        "This is an automated message. Please do not reply to this e-mail"
        + (f"; for questions write to {reply_to}" if reply_to else "") + ".\n"
        "If you were not expecting it, you can safely ignore it.\n")
    return SUBJECT, text, html


def build(settings, to, partner_name, link, expires_label, resign_note=None):
    subject, text, html = content(settings, partner_name, link, expires_label, resign_note)
    msg = EmailMessage()
    msg["Subject"] = header_safe(subject)
    msg["From"] = _sender(settings.smtp) or header_safe(settings.smtp.get("from"))
    msg["To"] = formataddr((header_safe(partner_name, 120), header_safe(to, 254)))
    reply_to = settings.smtp.get("reply_to") or ""
    if valid_email(reply_to):
        msg["Reply-To"] = reply_to
    msg["Message-ID"] = make_msgid(domain="dupoin-dw-calculator")
    msg["Auto-Submitted"] = "auto-generated"
    msg["X-Auto-Response-Suppress"] = "All"
    msg.set_content(text)
    msg.add_alternative(html, subtype="html")
    return msg


def scrub(text, *secrets):
    """Never let a password or signing token reach a note, the audit log or a log."""
    out = text or ""
    for s in secrets:
        if s and len(s) >= 4:
            out = out.replace(s, "***")
    return re.sub(r"/sign/[A-Za-z0-9_-]{16,}", "/sign/***", out)


def send_signing_link(settings, to, partner_name, link, expires_label, resign_note=None):
    """Returns (sent: bool, note: str). Never raises: a mail problem must not
    undo an approval; the admin sees the note and can resend or copy the link."""
    if not settings.smtp_ready:
        return False, NOT_CONFIGURED
    if not to:
        return False, "Email not sent: the request has no partner e-mail"
    if not valid_email(to):
        return False, "Email not sent: the partner e-mail address is not valid"
    s = settings.smtp
    if _sender(s) is None:
        return False, "Email not sent: DW_SMTP_FROM is not a valid e-mail address"
    msg = build(settings, to, partner_name, link, expires_label, resign_note)
    try:
        if s["tls"] == "ssl":
            server = smtplib.SMTP_SSL(s["host"], s["port"], timeout=30,
                                      context=ssl.create_default_context())
        else:
            server = smtplib.SMTP(s["host"], s["port"], timeout=30)
        with server:
            if s["tls"] == "starttls":
                server.starttls(context=ssl.create_default_context())
            if s["user"]:
                server.login(s["user"], s["password"])
            server.send_message(msg)
    except Exception as exc:          # noqa: BLE001 - reported to the admin, not raised
        token = link.rsplit("/", 1)[-1] if link else ""
        detail = scrub(f"{type(exc).__name__}: {str(exc)[:200]}", s.get("password"), token)
        return False, f"Email not sent: {detail}"
    return True, f"Email sent to {to}"
