"""Partner Approval (Beta) -- DPM only, test-only.

Self-contained so it can later move to FinanceOS: two Flask blueprints, its own
SQLite file, its own storage folder, its own templates. The host app provides
only the signed-in user (request.user) and the company check.

    /approval/...      DPM users: submit a Sales Partner Request, see their own
    /approval/admin/.. admins: approve / reject, review signatures, complete
    /sign/<token>      public: the partner reads and signs the agreement

Wire-up from app.py:  partner_approval.register(app)
"""
import hashlib
import hmac
import re
import secrets
import time
from pathlib import Path

from flask import (Blueprint, abort, current_app, flash, g, jsonify, redirect, render_template,
                   request, send_file, url_for)
from werkzeug.exceptions import HTTPException

from . import config, docs, forms, mailer
from .store import SIGNABLE, STATUS, TRANSITIONS, Store, ref, today_wib, wib

HERE = Path(__file__).resolve().parent
bp = Blueprint("partner_approval", __name__, url_prefix="/approval",
               template_folder="templates")
sign_bp = Blueprint("partner_sign", __name__, url_prefix="/sign",
                    template_folder="templates", static_folder="static",
                    static_url_path="/static")

SETTINGS = None
STORE = None
BETA = "Beta"

# Brute-force guard for the public /sign/* pages (shared by all workers via
# SQLite). A link has 256 bits of randomness, so this is defence in depth.
BAD_TOKEN_LIMIT = 20          # unknown links per IP ...
POST_LIMIT = 30               # sign/upload posts per IP ...
LIMIT_WINDOW = 15 * 60        # ... per 15 minutes
RESEND_GAP = 60               # seconds between two e-mails of the same request
LONG_DATE = "%d %B %Y, %H:%M WIB"
SIGN_FORM_MAX = 3 * 1024 * 1024               # drawn signature data URL
SIGN_UPLOAD_MAX = docs.MAX_UPLOAD + 1024 * 1024
_SUBMIT_KEY = re.compile(r"[A-Za-z0-9_-]{16,64}")


def register(app):
    global SETTINGS, STORE
    SETTINGS = config.load()
    STORE = Store(SETTINGS.db)
    app.register_blueprint(bp)
    app.register_blueprint(sign_bp)
    app.extensions["partner_approval"] = {"settings": SETTINGS, "store": STORE}


def reload_settings():
    """Tests change env vars (SMTP, link days) after import."""
    global SETTINGS
    SETTINGS = config.load()
    return SETTINGS


# ---------------------------------------------------------------- helpers
def _user():
    return getattr(request, "user", None)


def _is_admin():
    u = _user()
    return bool(u and u["is_admin"])


def _csrf_token():
    u = _user()
    name = (u["username"] if u else "").encode()
    key = (current_app.secret_key or "").encode()
    return hmac.new(key, b"partner-approval-csrf:" + name, hashlib.sha256).hexdigest()


@bp.context_processor
def _ctx():
    return {"csrf_token": _csrf_token, "STATUS": STATUS, "BETA": BETA, "wib": wib,
            "is_admin": _is_admin(), "types": forms.REQUEST_TYPES}


def _is_admin_route():
    rule = request.url_rule.rule if request.url_rule is not None else request.path
    return rule.startswith("/approval/admin") or request.path.startswith("/approval/admin")


@bp.before_request
def _guard():
    # app.before_request already required a signed-in DPM user for /approval.
    if _user() is None:
        abort(401)
    if _is_admin_route() and not _is_admin():
        abort(403, description="Only an admin can review partner requests.")
    if request.method == "POST":
        sent = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token") or ""
        if not hmac.compare_digest(sent.encode(), _csrf_token().encode()):
            abort(400, description="The form expired. Reload the page and try again.")


def _files_dir(rid):
    d = SETTINGS.files / f"{int(rid):06d}"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _rel(path):
    return str(Path(path).resolve().relative_to(SETTINGS.files.resolve()))


def _abs(rel):
    if not rel:
        return None
    p = (SETTINGS.files / rel).resolve()
    if SETTINGS.files.resolve() not in p.parents:
        return None
    return p if p.is_file() else None


def _link(token):
    return f"{SETTINGS.base_url}/sign/{token}"


def _get_or_404(rid):
    req = STORE.get(rid)
    if req is None:
        abort(404)
    return req


def _agreement_applies(req):
    return req["request_type"] in forms.AGREEMENT_TYPES


def _issue_agreement(req, actor, resign_note=None, regenerate=True):
    """Fill + convert the agreement, create a fresh signing link, e-mail it.

    Never raises for an expected problem (no template, no LibreOffice, no SMTP):
    the reason is stored in agreement_note / email_note and shown to the admin.
    Returns the new token, or None when no link could be made.
    """
    rid, data = req["id"], req["data"]
    template = docs.template_for(SETTINGS.templates, data.get("role"))
    if template is None:
        note = f"Agreement template for {data.get('role')} not ready yet"
        STORE.update(rid, agreement_note=note)
        STORE.log(rid, actor, note)
        return None
    pdf = _abs(req.get("agreement_pdf"))
    if regenerate or pdf is None:
        out = _files_dir(rid)
        docx_path = out / "agreement.docx"
        docs.fill_agreement(template, data, docx_path)
        STORE.update(rid, agreement_docx=_rel(docx_path))
        try:
            pdf = docs.docx_to_pdf(docx_path, out / "agreement.pdf", SETTINGS.soffice)
        except (docs.ConversionUnavailable, docs.ConversionFailed) as exc:
            STORE.update(rid, agreement_note=str(exc), agreement_pdf=None)
            STORE.log(rid, actor, f"Agreement PDF not created: {exc}")
            return None
        STORE.update(rid, agreement_pdf=_rel(pdf))
        docs.pdf_previews(pdf, _files_dir(rid) / "preview")
    token = STORE.new_token(rid, SETTINGS.link_days)
    tok = STORE.token(token)
    STORE.update(rid, agreement_note=None)
    STORE.log(rid, actor, f"Signing link created ({template.name}), valid until "
                          f"{wib(tok['expires_at'])}")
    _email_link(rid, data, token, tok["expires_at"], actor, resign_note)
    return token


def _email_link(rid, data, token, expires_at, actor, resign_note=None, prefix=""):
    """Send (or resend) the signing e-mail; the result goes to email_note + audit."""
    sent, note = mailer.send_signing_link(SETTINGS, data.get("partner_email"),
                                          data.get("partner_name"), _link(token),
                                          wib(expires_at, LONG_DATE), resign_note)
    note = mailer.scrub(note, SETTINGS.smtp.get("password"), token)
    STORE.update(rid, email_note=note, email_at=time.time())
    STORE.log(rid, actor, prefix + note)
    return sent, note


def _open_link(rid, req=None):
    """The request's current signing link when it can still be used, else None."""
    req = req or STORE.get(rid)
    tok = STORE.current_token(rid)
    if (req is None or tok is None or req["status"] not in SIGNABLE or tok["used_at"]
            or tok["revoked_at"] or tok["expires_at"] < time.time()):
        return None
    return tok


# ----------------------------------------------------------- user pages
@bp.get("")
@bp.get("/")
def index():
    request.company = "dpm"
    mine = STORE.list(user=_user()["username"])
    return render_template("approval/index.html", requests=mine)


def _submit_key():
    key = request.form.get("submit_key") or ""
    return key if _SUBMIT_KEY.fullmatch(key) else None


@bp.get("/new")
def new():
    request.company = "dpm"
    return render_template("approval/form.html", f=forms, values={
        "effective_date": today_wib().isoformat()}, errors={},
        submit_key=secrets.token_urlsafe(18))


@bp.post("/preview")
def preview():
    data, errors = forms.validate(request.form)
    if errors:
        return jsonify(ok=False, errors=errors), 400
    return jsonify(ok=True, summary=forms.summary(data), details=forms.details(data))


@bp.post("/new")
def submit():
    request.company = "dpm"
    data, errors = forms.validate(request.form)
    key = _submit_key() or secrets.token_urlsafe(18)
    if errors:
        return render_template("approval/form.html", f=forms, values=request.form,
                               errors=errors, submit_key=key), 400
    if request.form.get("confirmed") != "1":
        # Without JavaScript: the confirmation is a page instead of a popup.
        return render_template("approval/confirm.html", summary=forms.summary(data),
                               details=forms.details(data), values=data, submit_key=key)
    # The same submit_key twice (double click, back + resubmit) = the same request.
    rid = STORE.create(data, _user()["username"], submit_key=_submit_key())
    flash(f"Request {ref(rid)} submitted. Status: Pending.")
    return redirect(url_for("partner_approval.mine", rid=rid))


@bp.get("/<int:rid>")
def mine(rid):
    request.company = "dpm"
    req = _get_or_404(rid)
    if req["created_by"] != _user()["username"] and not _is_admin():
        abort(404)
    return render_template("approval/mine.html", req=req, summary=forms.summary(req["data"]),
                           details=forms.details(req["data"]), audit=STORE.audit(rid))


# ----------------------------------------------------------- admin pages
@bp.get("/admin")
def admin_list():
    request.company = "dpm"
    status = request.args.get("status") or ""
    if status and status not in STATUS:
        status = ""
    return render_template("approval/admin_list.html", requests=STORE.list(status=status or None),
                           status=status, counts=STORE.counts(), types=forms.REQUEST_TYPES)


@bp.get("/admin/<int:rid>")
def admin_detail(rid):
    request.company = "dpm"
    req = _get_or_404(rid)
    tok = STORE.current_token(rid)
    link = None
    if tok and req["status"] in SIGNABLE and not tok["used_at"]:
        link = {"url": _link(tok["token"]), "expires": wib(tok["expires_at"]),
                "expired": tok["expires_at"] < time.time() or bool(tok["revoked_at"]),
                "opened": wib(tok["opened_at"]) if tok["opened_at"] else None}
    template = docs.template_for(SETTINGS.templates, req["data"].get("role"))
    note = req.get("email_note") or ""
    return render_template(
        "approval/admin_detail.html", req=req, summary=forms.summary(req["data"]),
        details=forms.details(req["data"]), email_ok=note.startswith("Email sent"),
        allowed=TRANSITIONS.get(req["status"], ()), audit=STORE.audit(rid), link=link, agreement=_agreement_applies(req),
        template=template.name if template else None, smtp_ready=SETTINGS.smtp_ready,
        files={k: bool(_abs(req.get(k))) for k in
               ("agreement_pdf", "agreement_docx", "signed_file", "signed_docx")},
        )


def _actor():
    return _user()["username"]


@bp.post("/admin/<int:rid>/approve")
def approve(rid):
    req = _get_or_404(rid)
    if req["status"] != "pending":
        flash(f"{req['ref']} is no longer Pending.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    if not _agreement_applies(req):
        if STORE.set_status(rid, "completed", _actor(),
                            "Approved (no agreement for this request type)", expect=("pending",)):
            flash(f"{req['ref']} approved.")
        else:
            flash(f"{req['ref']} was already handled by someone else.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    if not STORE.set_status(rid, "approved", _actor(), "Approved", expect=("pending",)):
        flash(f"{req['ref']} was already handled by someone else.")     # double click / race
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    token = _issue_agreement(STORE.get(rid), _actor())
    after = STORE.get(rid)
    flash(f"{req['ref']} approved. " + ((after["email_note"] if token else after["agreement_note"])
                                        or ""))
    return redirect(url_for("partner_approval.admin_detail", rid=rid))


@bp.post("/admin/<int:rid>/reject")
def reject(rid):
    req = _get_or_404(rid)
    reason = " ".join((request.form.get("reason") or "").split())[:1000]
    if not reason:
        flash("Write the reason for rejecting.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    if not STORE.set_status(rid, "rejected", _actor(), f"Rejected: {reason}",
                            expect=("pending", "approved", "signed", "resign"),
                            reject_reason=reason):
        flash(f"{req['ref']} cannot be rejected now.")
    else:
        STORE.revoke(rid)                       # close any open link
        flash(f"{req['ref']} rejected.")
    return redirect(url_for("partner_approval.admin_detail", rid=rid))


@bp.post("/admin/<int:rid>/retry")
def retry(rid):
    req = _get_or_404(rid)
    if req["status"] not in SIGNABLE or not _agreement_applies(req):
        abort(409)
    if _open_link(rid, req):
        # A double click must not make (and e-mail) a second link.
        flash("A valid signing link already exists. Use \u201cResend email\u201d to send it again.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    token = _issue_agreement(req, _actor())
    after = STORE.get(rid)
    flash(("Signing link created. " + (after["email_note"] or "")) if token
          else (after["agreement_note"] or "No signing link was created."))
    return redirect(url_for("partner_approval.admin_detail", rid=rid))


@bp.post("/admin/<int:rid>/resend")
def resend(rid):
    """E-mail the current signing link again (e.g. after SMTP failed). No new link."""
    req = _get_or_404(rid)
    tok = _open_link(rid, req)
    if tok is None:
        flash("There is no valid signing link to send. Create a new link first.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    if (req.get("email_at") and time.time() - req["email_at"] < RESEND_GAP
            and (req.get("email_note") or "").startswith("Email sent")):
        flash("The e-mail was sent less than a minute ago. Please wait before sending it again.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    _sent, note = _email_link(rid, req["data"], tok["token"], tok["expires_at"], _actor(),
                              req.get("resign_note") if req["status"] == "resign" else None,
                              prefix="Resend: ")
    flash(note)
    return redirect(url_for("partner_approval.admin_detail", rid=rid))


@bp.post("/admin/<int:rid>/resign")
def resign(rid):
    req = _get_or_404(rid)
    note = " ".join((request.form.get("note") or "").split())[:1000]
    if not _agreement_applies(req) or req["status"] not in ("approved", "signed", "resign"):
        abort(409)
    tok = STORE.current_token(rid)
    if req["status"] == "resign" and tok and time.time() - tok["created_at"] < 10:
        flash("Re-sign was just requested; the new link is below.")       # double click
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    if not STORE.set_status(rid, "resign", _actor(),
                            "Re-sign requested" + (f": {note}" if note else ""),
                            expect=("approved", "signed", "resign"), resign_note=note or None):
        flash(f"{req['ref']} cannot be sent for re-sign now.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    token = _issue_agreement(STORE.get(rid), _actor(), resign_note=note or None, regenerate=False)
    after = STORE.get(rid)
    flash("Re-sign requested. " + (after["email_note"] if token else after["agreement_note"] or ""))
    return redirect(url_for("partner_approval.admin_detail", rid=rid))


@bp.post("/admin/<int:rid>/complete")
def complete(rid):
    req = _get_or_404(rid)
    if not STORE.set_status(rid, "completed", _actor(), "Marked completed", expect=("signed",)):
        flash(f"{req['ref']} can be completed only after it was signed.")
    else:
        flash(f"{req['ref']} completed.")
    return redirect(url_for("partner_approval.admin_detail", rid=rid))


FILE_KINDS = {"agreement_pdf": "agreement", "agreement_docx": "agreement",
              "signed_file": "signed", "signed_docx": "signed"}


@bp.get("/admin/<int:rid>/file/<kind>")
def admin_file(rid, kind):
    req = _get_or_404(rid)
    if kind not in FILE_KINDS:
        abort(404)
    path = _abs(req.get(kind))
    if path is None:
        abort(404)
    name = f"{req['ref']}-{FILE_KINDS[kind]}{path.suffix}"
    mimetype = MIMETYPES.get(path.suffix.lower(), "application/octet-stream")
    inline = request.args.get("inline") == "1" and path.suffix.lower() in (".pdf", ".png", ".jpg")
    if not inline and kind.startswith("signed"):
        STORE.log(rid, _actor(), f"Downloaded {name}")
    resp = send_file(path, mimetype=mimetype, as_attachment=not inline, download_name=name)
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Cache-Control"] = "private, no-store"
    if mimetype.startswith("image/"):
        resp.headers["Content-Security-Policy"] = "default-src 'none'; sandbox"
    return resp


MIMETYPES = {".pdf": "application/pdf", ".png": "image/png", ".jpg": "image/jpeg",
             ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}


# ------------------------------------------------------------ signing page
def _sign_csrf(token):
    """Form token of the public sign/upload forms, bound to the link."""
    key = (current_app.secret_key or "").encode()
    return hmac.new(key, b"partner-sign-csrf:" + token.encode(), hashlib.sha256).hexdigest()


def _csp_nonce():
    if "pa_nonce" not in g:
        g.pa_nonce = secrets.token_urlsafe(16)
    return g.pa_nonce


@sign_bp.context_processor
def _sign_ctx():
    return {"csp_nonce": _csp_nonce, "sign_csrf": _sign_csrf, "wib": wib}


def _limited():
    return render_template("approval/sign_closed.html", state="limited"), 429


@sign_bp.before_request
def _sign_guard():
    if request.endpoint == "partner_sign.static":
        return None
    # Smaller body limits than the app-wide 1 GB upload limit.
    if request.endpoint == "partner_sign.sign_upload":
        request.max_content_length = SIGN_UPLOAD_MAX
    else:
        request.max_content_length = SIGN_FORM_MAX + 64 * 1024
    request.max_form_memory_size = SIGN_FORM_MAX
    ip = _client_ip()
    if STORE.hits("bad:" + ip, LIMIT_WINDOW) >= BAD_TOKEN_LIMIT:
        return _limited()
    if request.method == "POST":
        if STORE.hits("post:" + ip, LIMIT_WINDOW) >= POST_LIMIT:
            return _limited()
        STORE.hit("post:" + ip)
    return None


def _no_cache(resp):
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    if request.endpoint in ("partner_sign.sign_pdf", "partner_sign.sign_preview"):
        # Shown inside our own signing page only.
        resp.headers["X-Frame-Options"] = "SAMEORIGIN"
        resp.headers["Content-Security-Policy"] = "frame-ancestors 'self'"
    elif request.endpoint != "partner_sign.static":
        resp.headers["X-Frame-Options"] = "DENY"
        resp.headers["Content-Security-Policy"] = (
            "default-src 'none'; "
            f"script-src 'nonce-{_csp_nonce()}'; "
            "style-src 'self' 'unsafe-inline'; img-src 'self' data:; font-src 'self'; "
            "frame-src 'self'; connect-src 'self'; form-action 'self'; base-uri 'none'; "
            "frame-ancestors 'none'")
        resp.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    return resp


@sign_bp.after_request
def _sign_headers(resp):
    return _no_cache(resp)


@sign_bp.errorhandler(Exception)
def _sign_crash(exc):
    """Log without the request path: the path is the partner's secret link."""
    if isinstance(exc, HTTPException):
        if exc.code == 413:
            return render_template("approval/sign_closed.html", state="too_large"), 413
        return exc
    current_app.logger.error("Partner sign page error on %s /sign/***: %s", request.method,
                             type(exc).__name__, exc_info=exc)
    return render_template("approval/sign_closed.html", state="error"), 500


def _token_state(token):
    """(state, token_row, request). state: ok | invalid | revoked | used | expired | closed"""
    row = STORE.token(token)
    if row is None:
        STORE.hit("bad:" + _client_ip())
        return "invalid", None, None
    req = STORE.get(row["request_id"])
    if row["used_at"]:
        return "used", row, req
    if row["revoked_at"]:
        return "revoked", row, req
    if row["expires_at"] < time.time():
        return "expired", row, req
    if req is None or req["status"] not in SIGNABLE:
        return "closed", row, req
    return "ok", row, req


STATE_CODES = {"invalid": 404, "revoked": 410, "used": 200, "expired": 410, "closed": 410}


def _closed_page(state):
    return render_template("approval/sign_closed.html", state=state), STATE_CODES[state]


def _client_ip():
    for h in ("CF-Connecting-IP", "X-Real-IP"):
        if request.headers.get(h):
            return request.headers[h].strip()[:64]
    fwd = request.headers.get("X-Forwarded-For")
    if fwd:
        return fwd.split(",")[0].strip()[:64]
    return (request.remote_addr or "")[:64]


def _previews(req):
    pdf = _abs(req.get("agreement_pdf"))
    folder = _files_dir(req["id"]) / "preview"
    n = docs.preview_count(folder)
    if not n and pdf is not None:
        n = docs.pdf_previews(pdf, folder)        # requests made before v1.34
    return n


def _render_sign(req, token, row, error=None, mode="draw", typed=None):
    return render_template("approval/sign.html", req=req, token=token,
                           expires=wib(row["expires_at"], LONG_DATE), error=error,
                           pages=_previews(req), mode=mode, typed=typed,
                           max_mb=docs.MAX_UPLOAD // (1024 * 1024))


@sign_bp.get("/<token>")
def sign_page(token):
    state, row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    if not _abs(req.get("agreement_pdf")):
        return _closed_page("closed")
    if STORE.mark_opened(token):
        STORE.log(req["id"], "partner", f"Signing page opened (IP {_client_ip()})")
    return _render_sign(req, token, row)


@sign_bp.get("/<token>/agreement.pdf")
def sign_pdf(token):
    state, _row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    path = _abs(req.get("agreement_pdf"))
    if path is None:
        abort(404)
    return send_file(path, mimetype="application/pdf",
                     as_attachment=request.args.get("download") == "1",
                     download_name=f"{req['ref']}-agreement.pdf")


@sign_bp.get("/<token>/page/<int:n>.png")
def sign_preview(token, n):
    state, _row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    if not 1 <= n <= docs.MAX_PREVIEW_PAGES:
        abort(404)
    path = _abs(f"{int(req['id']):06d}/preview/page-{n}.png")
    if path is None:
        abort(404)
    return send_file(path, mimetype="image/png")


def _sign_error(req, token, row, message, code=400, **kw):
    return _render_sign(req, token, row, error=message, **kw), code


def _csrf_ok(token):
    sent = request.form.get("csrf") or ""
    return hmac.compare_digest(sent.encode(), _sign_csrf(token).encode())


def _signed(rid, ok):
    """After the link was used: the request may have been rejected meanwhile."""
    if ok:
        return render_template("approval/sign_done.html", req=STORE.get(rid))
    STORE.log(rid, "system", "A signature arrived after the request was closed; it was not accepted")
    return _closed_page("closed")


@sign_bp.post("/<token>/sign")
def sign_submit(token):
    state, row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    mode = request.form.get("mode")
    typed = (request.form.get("typed_name") or "")[:200]
    if not _csrf_ok(token):
        return _sign_error(req, token, row, "This page has expired. Please reload it and sign again.",
                           mode=mode, typed=typed)
    if request.form.get("agree") != "1":
        return _sign_error(req, token, row, "Please tick the box to confirm that you agree.",
                           mode=mode, typed=typed)
    try:
        if mode == "draw":
            png = docs.drawn_signature(request.form.get("signature") or "")
        elif mode == "typed":
            png = docs.typed_signature(forms._clean(typed, False))
        else:
            return _sign_error(req, token, row, "Choose how you want to sign.")
    except docs.SignatureError as exc:
        return _sign_error(req, token, row, str(exc), mode=mode, typed=typed)
    if not STORE.use_token(token):
        return _closed_page("used")
    rid = req["id"]
    out = _files_dir(rid)
    (out / "signature.png").write_bytes(png)
    sign_date = today_wib()
    signed_docx = out / "signed.docx"
    note = None
    signed_file = None
    try:
        template = docs.template_for(SETTINGS.templates, req["data"].get("role"))
        docs.fill_agreement(template, req["data"], signed_docx, sign_date=sign_date,
                            signature_png=png)
        signed_file = _rel(docs.docx_to_pdf(signed_docx, out / "signed.pdf", SETTINGS.soffice))
    except (docs.ConversionUnavailable, docs.ConversionFailed) as exc:
        note = f"Signed PDF not created: {exc}"
    except Exception as exc:  # noqa: BLE001 - the partner must still get through
        note = f"Signed document not created: {type(exc).__name__}: {exc}"
    ok = STORE.set_status(rid, "signed", "partner",
                          f"Signed on screen ({'drawn' if mode == 'draw' else 'typed name'}) "
                          f"from IP {_client_ip()}",
                          expect=SIGNABLE, signed_kind="drawn" if mode == "draw" else "typed",
                          signed_file=signed_file,
                          signed_docx=_rel(signed_docx) if signed_docx.is_file() else None,
                          signed_at=time.time(), signer_ip=_client_ip(),
                          signer_ua=(request.headers.get("User-Agent") or "")[:400],
                          agreement_note=note)
    if note and ok:
        STORE.log(rid, "system", note)
    return _signed(rid, ok)


@sign_bp.post("/<token>/upload")
def sign_upload(token):
    state, row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    if not _csrf_ok(token):
        return _sign_error(req, token, row, "This page has expired. Please reload it and upload again.",
                           mode="upload")
    f = request.files.get("file")
    if not f or not f.filename:
        return _sign_error(req, token, row, "Choose the signed file to upload.", mode="upload")
    blob = f.stream.read(docs.MAX_UPLOAD + 1)
    if len(blob) > docs.MAX_UPLOAD:
        return _sign_error(req, token, row,
                           f"The file is larger than {docs.MAX_UPLOAD // (1024 * 1024)} MB.", 413,
                           mode="upload")
    ext = docs.check_upload(blob)
    if ext is None:
        return _sign_error(req, token, row, "Upload a PDF, JPG or PNG file.", mode="upload")
    if not STORE.use_token(token):
        return _closed_page("used")
    rid = req["id"]
    path = _files_dir(rid) / f"signed_upload.{ext}"     # our name, never the partner's
    path.write_bytes(blob)
    ok = STORE.set_status(rid, "signed", "partner", f"Signed copy uploaded ({ext.upper()}, "
                          f"{len(blob):,} bytes) from IP {_client_ip()}", expect=SIGNABLE,
                          signed_kind="upload", signed_file=_rel(path), signed_docx=None,
                          signed_at=time.time(), signer_ip=_client_ip(),
                          signer_ua=(request.headers.get("User-Agent") or "")[:400])
    return _signed(rid, ok)
