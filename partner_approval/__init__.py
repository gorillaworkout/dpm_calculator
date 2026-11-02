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
import os
from pathlib import Path

from flask import (Blueprint, abort, current_app, flash, jsonify, redirect, render_template,
                   request, send_file, url_for)

from . import config, docs, forms, mailer
from .store import SIGNABLE, STATUS, Store, ref, today_wib, wib

HERE = Path(__file__).resolve().parent
bp = Blueprint("partner_approval", __name__, url_prefix="/approval",
               template_folder="templates")
sign_bp = Blueprint("partner_sign", __name__, url_prefix="/sign",
                    template_folder="templates", static_folder="static",
                    static_url_path="/static")

SETTINGS = None
STORE = None
BETA = "Beta"


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


@bp.before_request
def _guard():
    # app.before_request already required a signed-in DPM user for /approval.
    if _user() is None:
        abort(401)
    if request.method == "POST":
        sent = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token") or ""
        if not hmac.compare_digest(sent, _csrf_token()):
            abort(400, description="The form expired. Reload the page and try again.")
    if request.path.startswith("/approval/admin") and not _is_admin():
        abort(403, description="Only an admin can review partner requests.")


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
    token = STORE.new_token(rid, SETTINGS.link_days)
    tok = STORE.token(token)
    STORE.update(rid, agreement_note=None)
    STORE.log(rid, actor, f"Signing link created ({template.name}), valid until "
                          f"{wib(tok['expires_at'])}")
    sent, note = mailer.send_signing_link(SETTINGS, data.get("partner_email"),
                                          data.get("partner_name"), _link(token),
                                          wib(tok["expires_at"]), resign_note)
    STORE.update(rid, email_note=note)
    STORE.log(rid, actor, note)
    return token


# ----------------------------------------------------------- user pages
@bp.get("")
@bp.get("/")
def index():
    request.company = "dpm"
    mine = STORE.list(user=_user()["username"])
    return render_template("approval/index.html", requests=mine)


@bp.get("/new")
def new():
    request.company = "dpm"
    return render_template("approval/form.html", f=forms, values={
        "effective_date": today_wib().isoformat()}, errors={})


@bp.post("/preview")
def preview():
    data, errors = forms.validate(request.form)
    if errors:
        return jsonify(ok=False, errors=errors), 400
    return jsonify(ok=True, summary=forms.summary(data))


@bp.post("/new")
def submit():
    request.company = "dpm"
    data, errors = forms.validate(request.form)
    if errors:
        return render_template("approval/form.html", f=forms, values=request.form,
                               errors=errors), 400
    if request.form.get("confirmed") != "1":
        # Without JavaScript: the confirmation is a page instead of a popup.
        return render_template("approval/confirm.html", summary=forms.summary(data),
                               values=data)
    rid = STORE.create(data, _user()["username"])
    flash(f"Request {ref(rid)} submitted. Status: Pending.")
    return redirect(url_for("partner_approval.mine", rid=rid))


@bp.get("/<int:rid>")
def mine(rid):
    request.company = "dpm"
    req = _get_or_404(rid)
    if req["created_by"] != _user()["username"] and not _is_admin():
        abort(404)
    return render_template("approval/mine.html", req=req, summary=forms.summary(req["data"]),
                           audit=STORE.audit(rid))


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
                "expired": tok["expires_at"] < __import__("time").time()}
    template = docs.template_for(SETTINGS.templates, req["data"].get("role"))
    return render_template(
        "approval/admin_detail.html", req=req, summary=forms.summary(req["data"]),
        audit=STORE.audit(rid), link=link, agreement=_agreement_applies(req),
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
        STORE.set_status(rid, "completed", _actor(), "Approved (no agreement for this request type)",
                         expect=("pending",))
        flash(f"{req['ref']} approved.")
        return redirect(url_for("partner_approval.admin_detail", rid=rid))
    if not STORE.set_status(rid, "approved", _actor(), "Approved", expect=("pending",)):
        abort(409)
    token = _issue_agreement(STORE.get(rid), _actor())
    after = STORE.get(rid)
    flash(f"{req['ref']} approved. " + (after["email_note"] if token else after["agreement_note"] or ""))
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
    token = _issue_agreement(req, _actor())
    after = STORE.get(rid)
    flash("Signing link created. " + after["email_note"] if token else after["agreement_note"])
    return redirect(url_for("partner_approval.admin_detail", rid=rid))


@bp.post("/admin/<int:rid>/resign")
def resign(rid):
    req = _get_or_404(rid)
    note = " ".join((request.form.get("note") or "").split())[:1000]
    if not _agreement_applies(req) or req["status"] not in ("approved", "signed", "resign"):
        abort(409)
    STORE.set_status(rid, "resign", _actor(), "Re-sign requested" + (f": {note}" if note else ""),
                     expect=("approved", "signed", "resign"), resign_note=note or None)
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
    inline = request.args.get("inline") == "1" and path.suffix in (".pdf", ".png", ".jpg")
    return send_file(path, as_attachment=not inline, download_name=name)


# ------------------------------------------------------------ signing page
def _no_cache(resp):
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    resp.headers["X-Frame-Options"] = "SAMEORIGIN"
    return resp


@sign_bp.after_request
def _sign_headers(resp):
    return _no_cache(resp)


def _token_state(token):
    """(state, token_row, request). state: ok | invalid | revoked | used | expired | closed"""
    import time
    row = STORE.token(token)
    if row is None:
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
            return request.headers[h][:64]
    fwd = request.headers.get("X-Forwarded-For")
    if fwd:
        return fwd.split(",")[0].strip()[:64]
    return (request.remote_addr or "")[:64]


@sign_bp.get("/<token>")
def sign_page(token):
    state, row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    if not _abs(req.get("agreement_pdf")):
        return _closed_page("closed")
    return render_template("approval/sign.html", req=req, token=token,
                           expires=wib(row["expires_at"]), error=None,
                           max_mb=docs.MAX_UPLOAD // (1024 * 1024))


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


def _sign_error(req, token, row, message, code=400):
    return render_template("approval/sign.html", req=req, token=token,
                           expires=wib(row["expires_at"]), error=message,
                           max_mb=docs.MAX_UPLOAD // (1024 * 1024)), code


@sign_bp.post("/<token>/sign")
def sign_submit(token):
    state, row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    if request.form.get("agree") != "1":
        return _sign_error(req, token, row, "Please tick 'I agree' to sign the agreement.")
    mode = request.form.get("mode")
    try:
        if mode == "draw":
            png = docs.drawn_signature(request.form.get("signature") or "")
        elif mode == "typed":
            png = docs.typed_signature(request.form.get("typed_name") or "")
        else:
            return _sign_error(req, token, row, "Choose how you want to sign.")
    except docs.SignatureError as exc:
        return _sign_error(req, token, row, str(exc))
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
    STORE.set_status(rid, "signed", "partner",
                     f"Signed on screen ({'drawn' if mode == 'draw' else 'typed name'})",
                     expect=SIGNABLE, signed_kind="drawn" if mode == "draw" else "typed",
                     signed_file=signed_file,
                     signed_docx=_rel(signed_docx) if signed_docx.is_file() else None,
                     signed_at=__import__("time").time(), signer_ip=_client_ip(),
                     signer_ua=(request.headers.get("User-Agent") or "")[:400],
                     agreement_note=note)
    if note:
        STORE.log(rid, "system", note)
    return render_template("approval/sign_done.html")


@sign_bp.post("/<token>/upload")
def sign_upload(token):
    state, row, req = _token_state(token)
    if state != "ok":
        return _closed_page(state)
    f = request.files.get("file")
    if not f or not f.filename:
        return _sign_error(req, token, row, "Choose the signed file to upload.")
    blob = f.stream.read(docs.MAX_UPLOAD + 1)
    if len(blob) > docs.MAX_UPLOAD:
        return _sign_error(req, token, row,
                           f"The file is larger than {docs.MAX_UPLOAD // (1024 * 1024)} MB.", 413)
    ext = docs.sniff_upload(blob[:16])
    if ext is None:
        return _sign_error(req, token, row, "Upload a PDF, JPG or PNG file.")
    if not STORE.use_token(token):
        return _closed_page("used")
    rid = req["id"]
    path = _files_dir(rid) / f"signed_upload.{ext}"
    path.write_bytes(blob)
    STORE.set_status(rid, "signed", "partner", f"Signed copy uploaded ({ext.upper()}, "
                     f"{len(blob):,} bytes)", expect=SIGNABLE, signed_kind="upload",
                     signed_file=_rel(path), signed_docx=None,
                     signed_at=__import__("time").time(), signer_ip=_client_ip(),
                     signer_ua=(request.headers.get("User-Agent") or "")[:400])
    return render_template("approval/sign_done.html")
