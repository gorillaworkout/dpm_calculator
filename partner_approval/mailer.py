"""Signing-link e-mail over SMTP. Nothing is sent when SMTP is not configured."""
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr, make_msgid

from markupsafe import escape

NOT_CONFIGURED = "Email not sent: SMTP not configured"


def build(settings, to, partner_name, link, expires_label, resign_note=None):
    msg = EmailMessage()
    msg["Subject"] = "Your Dupoin partnership agreement is ready to sign"
    msg["From"] = settings.smtp["from"]
    msg["To"] = formataddr((partner_name or "", to))
    msg["Message-ID"] = make_msgid(domain="dw-calculator")
    extra = f"\nNote from Dupoin: {resign_note}\n" if resign_note else ""
    text = (
        f"Dear {partner_name or 'Partner'},\n\n"
        "Your partnership agreement with Dupoin Markets Ltd is ready.\n"
        f"{extra}\n"
        "Please open the link below to read the agreement and sign it. You can sign "
        "on screen, or download the PDF, sign it by hand and upload it.\n\n"
        f"{link}\n\n"
        f"The link is personal and can be used once. It expires on {expires_label}.\n\n"
        "If you were not expecting this e-mail, please ignore it.\n\n"
        "Dupoin Markets Ltd\n")
    note_html = (f"<p><b>Note from Dupoin:</b> {escape(resign_note)}</p>" if resign_note else "")
    html = (
        "<!doctype html><html><body style=\"font-family:Arial,Helvetica,sans-serif;"
        "color:#1b2333;line-height:1.5\">"
        f"<p>Dear {escape(partner_name or 'Partner')},</p>"
        "<p>Your partnership agreement with Dupoin Markets Ltd is ready.</p>"
        f"{note_html}"
        "<p>Please open the link below to read the agreement and sign it. You can sign on "
        "screen, or download the PDF, sign it by hand and upload it.</p>"
        f"<p><a href=\"{escape(link)}\" style=\"display:inline-block;padding:10px 18px;"
        "background:#2f5bea;color:#fff;border-radius:6px;text-decoration:none\">"
        "Review and sign the agreement</a></p>"
        f"<p style=\"font-size:13px;color:#555\">Or copy this link: {escape(link)}<br>"
        f"The link is personal and can be used once. It expires on {escape(expires_label)}.</p>"
        "<p style=\"font-size:13px;color:#555\">If you were not expecting this e-mail, "
        "please ignore it.</p><p>Dupoin Markets Ltd</p></body></html>")
    msg.set_content(text)
    msg.add_alternative(html, subtype="html")
    return msg


def send_signing_link(settings, to, partner_name, link, expires_label, resign_note=None):
    """Returns (sent: bool, note: str). Never raises: a mail problem must not
    undo an approval; the admin sees the note and can copy the link instead."""
    if not settings.smtp_ready:
        return False, NOT_CONFIGURED
    if not to:
        return False, "Email not sent: the request has no partner e-mail"
    s = settings.smtp
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
        return False, f"Email not sent: {type(exc).__name__}: {str(exc)[:200]}"
    return True, f"Email sent to {to}"
