"""DPM Invitation: one public RSVP link per event (name, email, attend, allergy).

Self-contained like Partner Approval: two Flask blueprints, its own SQLite file,
its own image folder, its own templates. The host app provides only the
signed-in user (request.user) and the DPM company check.

    /invitation/...    DPM users: events, responses, export, delete a response
                       admins also: create / edit / open-close / delete an event
    /i/<slug>          public (no login): the invitation image and the form

Wire-up from app.py:  invitation.register(app)
"""
import hashlib
import hmac
import os
import re
import secrets
import time
from pathlib import Path

from flask import (Blueprint, Response, abort, current_app, flash, g, redirect, render_template,
                   request, send_file, url_for)
from werkzeug.exceptions import HTTPException

from . import config, export, images, rules
from .store import FILTERS, Store, is_open

bp = Blueprint("invitation", __name__, url_prefix="/invitation", template_folder="templates")
public_bp = Blueprint("invitation_public", __name__, url_prefix="/i", template_folder="templates")

SETTINGS = None
STORE = None

IP_LIMIT, IP_WINDOW = 60, 10 * 60            # submissions per IP per 10 minutes
EMAIL_LIMIT, EMAIL_WINDOW = 5, 60 * 60       # submissions per event + email per hour
PUBLIC_FORM_MAX = 64 * 1024                  # the RSVP form is a few short fields
UPLOAD_MAX = images.MAX_BYTES + 512 * 1024   # 8 MB image + the other form fields
_KEY = re.compile(images.KEY_RE)


def register(app):
    global SETTINGS, STORE
    SETTINGS = config.load()
    STORE = Store(SETTINGS.db)
    app.register_blueprint(bp)
    app.register_blueprint(public_bp)
    app.extensions["invitation"] = {"settings": SETTINGS, "store": STORE}


# ---------------------------------------------------------------- helpers
def _user():
    return getattr(request, "user", None)


def _is_admin():
    u = _user()
    return bool(u and u["is_admin"])


def _can_manage():
    """Anyone who can open the Invitation menu (DPM access) manages events.
    Only deleting a whole event stays admin-only (it removes every response)."""
    import auth
    u = _user()
    return bool(u and (u["is_admin"] or auth.has_company(u, "dpm")))


def _username():
    u = _user()
    return u["username"] if u else ""


def _csrf_token():
    key = (current_app.secret_key or "").encode()
    return hmac.new(key, b"invitation-csrf:" + _username().encode(), hashlib.sha256).hexdigest()


def _public_csrf(slug):
    key = (current_app.secret_key or "").encode()
    return hmac.new(key, b"invitation-public-csrf:" + slug.encode(), hashlib.sha256).hexdigest()


def _client_ip():
    # The same header order as Partner Approval (Cloudflare, then nginx).
    from partner_approval import _client_ip as pa_ip
    return pa_ip()


def public_link(event):
    return f"{SETTINGS.base_url}/i/{event['slug']}"


def _image_path(key):
    if not key or not _KEY.match(key):
        return None
    p = (SETTINGS.files / key).resolve()
    if SETTINGS.files.resolve() not in p.parents or not p.is_file():
        return None
    return p


def _save_image(blob):
    data, info = images.process(blob)
    SETTINGS.files.mkdir(parents=True, exist_ok=True)
    final = SETTINGS.files / info["key"]
    tmp = final.with_name(f".{final.name}.{secrets.token_hex(4)}.tmp")
    with open(tmp, "wb") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, final)
    return info


def _remove_image(key):
    p = _image_path(key)
    if p is None:
        return
    try:
        p.unlink()
    except OSError as exc:  # not fatal: the row is already gone or points elsewhere
        current_app.logger.warning("Invitation image %s not removed: %s", key, exc)


def _event_or_404(eid):
    ev = STORE.event(eid)
    if ev is None:
        abort(404, description="Invitation not found.")
    return ev


def _decorate(ev, now=None):
    now = now or time.time()
    ev = dict(ev)
    ev["is_open"] = is_open(ev, now)
    ev["auto_closed"] = ev["status"] == "open" and not ev["is_open"]
    ev["link"] = public_link(ev)
    ev["closes_label"] = rules.format_wib(ev.get("closes_at"))
    return ev


# ============================================================ admin pages
@bp.context_processor
def _ctx():
    return {"csrf_token": _csrf_token, "is_admin": _is_admin(), "can_manage": _can_manage(),
            "format_wib": rules.format_wib,
            "to_wib_input": rules.to_wib_input, "max_mb": images.MAX_BYTES // (1024 * 1024)}


# Create / edit / upload / open-close: every user with Invitation (DPM) access.
_MANAGE = {"invitation.create", "invitation.edit", "invitation.set_status"}
# Deleting a whole event (and all its responses) is destructive: admins only.
_ADMIN_ONLY = {"invitation.delete_event"}


@bp.before_request
def _guard():
    # app.before_request already required a signed-in DPM user for /invitation.
    if _user() is None:
        abort(401)
    if not _can_manage():
        abort(403, description="Your account does not have access to Invitation.")
    if request.endpoint in _ADMIN_ONLY and not _is_admin():
        abort(403, description="Only an admin can delete an invitation event.")
    if request.method == "POST":
        request.max_content_length = UPLOAD_MAX
        sent = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token") or ""
        if not hmac.compare_digest(sent.encode(), _csrf_token().encode()):
            abort(400, description="The form expired. Reload the page and try again.")


@bp.after_request
def _admin_headers(resp):
    # Responses hold personal data (emails, allergies).
    resp.headers["Cache-Control"] = "no-store"
    return resp


def _read_event_form(ev=None):
    """Name, slug and closing time from the create/edit form. Returns (fields, errors)."""
    errors = {}
    name, err = rules.parse_event_name(request.form.get("name"))
    if err:
        errors["name"] = err
    slug = (request.form.get("slug") or "").strip().lower() or rules.suggest_slug(name or "")
    err = rules.slug_issue(slug)
    if err:
        errors["slug"] = err
    elif STORE.slug_taken(slug, ev["id"] if ev else None):
        errors["slug"] = "That link is already used by another invitation."
    raw = request.form.get("closes_at") or ""
    if request.form.get("clear_closes_at"):
        raw = ""
    closes_at, err = None, None
    if ev is not None and raw and raw == rules.to_wib_input(ev.get("closes_at")):
        closes_at = ev["closes_at"]           # unchanged (may already be in the past)
    else:
        closes_at, err = rules.parse_closes_at(raw)
    if err:
        errors["closes_at"] = err
    return {"name": name, "slug": slug, "closes_at": closes_at}, errors


@bp.get("")
def index():
    now = time.time()
    events = [_decorate(e, now) for e in STORE.events()]
    return render_template("invitation/list.html", events=events, errors={}, form={})


@bp.post("/new")
def create():
    fields, errors = _read_event_form()
    upload = request.files.get("image")
    blob = upload.read(images.MAX_BYTES + 1) if upload and upload.filename else b""
    info = None
    if not errors:
        try:
            info = _save_image(blob)
        except images.ImageError as exc:
            errors["image"] = str(exc)
    elif not blob:
        errors["image"] = "Choose an image file."
    if errors:
        events = [_decorate(e) for e in STORE.events()]
        return render_template("invitation/list.html", events=events, errors=errors,
                               form=request.form, open_create=True), 422
    try:
        eid = STORE.create_event(fields["name"], fields["slug"], fields["closes_at"], info,
                                 _username())
    except Exception:
        _remove_image(info["key"])
        raise
    flash(f"Invitation “{fields['name']}” created. Share the public link below.")
    return redirect(url_for("invitation.detail", eid=eid), 303)


@bp.get("/<int:eid>")
def detail(eid):
    ev = _decorate(_event_or_404(eid))
    q = (request.args.get("q") or "").strip()[:120]
    filt = request.args.get("filter") or "all"
    if filt not in FILTERS:
        filt = "all"
    return render_template("invitation/detail.html", ev=ev, counts=STORE.counts(eid),
                           responses=STORE.responses(eid, q, filt), q=q, filt=filt,
                           errors={}, form={})


@bp.post("/<int:eid>/edit")
def edit(eid):
    ev = _event_or_404(eid)
    fields, errors = _read_event_form(ev)
    upload = request.files.get("image")
    blob = upload.read(images.MAX_BYTES + 1) if upload and upload.filename else b""
    info = None
    if blob and not errors:
        try:
            info = _save_image(blob)
        except images.ImageError as exc:
            errors["image"] = str(exc)
    if errors:
        dev = _decorate(ev)
        return render_template("invitation/detail.html", ev=dev, counts=STORE.counts(eid),
                               responses=STORE.responses(eid), q="", filt="all",
                               errors=errors, form=request.form, open_settings=True), 422
    changes = {"name": fields["name"], "slug": fields["slug"], "closes_at": fields["closes_at"]}
    if info:
        changes.update(image_key=info["key"], image_mime=info["mime"],
                       image_width=info["width"], image_height=info["height"])
    STORE.update_event(eid, _username(), **changes)
    if info:
        _remove_image(ev["image_key"])
    msg = "Invitation saved."
    if fields["slug"] != ev["slug"]:
        msg += " The link changed: the old link no longer works, share the new one."
    flash(msg)
    return redirect(url_for("invitation.detail", eid=eid), 303)


@bp.post("/<int:eid>/status")
def set_status(eid):
    ev = _event_or_404(eid)
    status = request.form.get("status")
    if status not in ("open", "closed"):
        abort(400, description="Unknown status.")
    STORE.update_event(eid, _username(), status=status)
    ev = _decorate(STORE.event(eid))
    if status == "open" and ev["auto_closed"]:
        flash(f"Registration is switched on, but it closed automatically at {ev['closes_label']}. "
              "Clear or move the closing time to reopen.")
    else:
        flash("Registration is open." if status == "open" else "Registration is closed.")
    return redirect(url_for("invitation.detail", eid=eid), 303)


@bp.post("/<int:eid>/delete")
def delete_event(eid):
    ev = _event_or_404(eid)
    if (request.form.get("confirm") or "").strip().lower() != ev["slug"]:
        flash("Not deleted: type the event link name exactly to confirm.")
        return redirect(url_for("invitation.detail", eid=eid), 303)
    STORE.delete_event(eid)
    _remove_image(ev["image_key"])
    flash(f"Invitation “{ev['name']}” and all its responses were deleted.")
    return redirect(url_for("invitation.index"), 303)


@bp.post("/<int:eid>/responses/<int:rid>/delete")
def delete_response(eid, rid):
    _event_or_404(eid)
    row = STORE.response(eid, rid)
    if row is None:
        abort(404, description="Response not found.")
    STORE.delete_response(eid, rid)
    flash(f"Response from {row['full_name']} deleted.")
    args = {k: v for k, v in (("q", request.form.get("q")), ("filter", request.form.get("filter"))) if v}
    return redirect(url_for("invitation.detail", eid=eid, **args), 303)


@bp.get("/<int:eid>/export")
def export_responses(eid):
    ev = _event_or_404(eid)
    fmt = request.args.get("format", "xlsx")
    if fmt not in ("xlsx", "csv"):
        abort(400, description="Choose xlsx or csv.")
    filt = request.args.get("filter") or "all"
    rows = STORE.responses(eid, request.args.get("q", "")[:120], filt if filt in FILTERS else "all")
    if fmt == "csv":
        body, mime = export.to_csv(rows), "text/csv; charset=utf-8"
    else:
        body = export.to_xlsx(rows, ev["name"])
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    name = export.filename(ev["slug"], fmt)
    return Response(body, mimetype=mime,
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


@bp.get("/<int:eid>/image")
def admin_image(eid):
    ev = _event_or_404(eid)
    path = _image_path(ev["image_key"])
    if path is None:
        abort(404)
    resp = send_file(path, mimetype=ev["image_mime"])
    resp.headers["X-Content-Type-Options"] = "nosniff"
    return resp


# ============================================================ public pages
def _nonce():
    if "inv_nonce" not in g:
        g.inv_nonce = secrets.token_urlsafe(16)
    return g.inv_nonce


@public_bp.context_processor
def _public_ctx():
    return {"csp_nonce": _nonce, "public_csrf": _public_csrf,
            "NAME_MAX": rules.NAME_MAX, "EMAIL_MAX": rules.EMAIL_MAX,
            "ALLERGY_MAX": rules.ALLERGY_MAX}


@public_bp.before_request
def _public_guard():
    request.max_content_length = PUBLIC_FORM_MAX
    request.max_form_memory_size = PUBLIC_FORM_MAX


@public_bp.after_request
def _public_headers(resp):
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    if request.endpoint != "invitation_public.image":
        resp.headers["Cache-Control"] = "no-store"
        resp.headers["X-Frame-Options"] = "DENY"
        resp.headers["Content-Security-Policy"] = (
            "default-src 'none'; "
            f"script-src 'nonce-{_nonce()}'; "
            "style-src 'unsafe-inline'; img-src 'self'; form-action 'self'; base-uri 'none'; "
            "frame-ancestors 'none'")
    return resp


@public_bp.errorhandler(HTTPException)
def _public_http_error(exc):
    if exc.code == 413:
        return render_template("invitation/public_message.html", ev=None, state="too_large"), 413
    if exc.code == 404:
        return render_template("invitation/public_message.html", ev=None, state="invalid"), 404
    return exc


def _public_event(slug):
    ev = STORE.event_by_slug((slug or "").lower())
    if ev is None:
        abort(404)
    return _decorate(ev)


def _render_form(ev, form=None, errors=None, notice=None, code=200):
    return render_template("invitation/public_form.html", ev=ev, form=form or {},
                           errors=errors or {}, notice=notice), code


def _closed(ev, code=200):
    return render_template("invitation/public_message.html", ev=ev, state="closed"), code


@public_bp.get("/<slug>")
def form(slug):
    ev = _public_event(slug)
    if not ev["is_open"]:
        return _closed(ev)
    return _render_form(ev)


@public_bp.post("/<slug>")
def submit(slug):
    ev = _public_event(slug)
    if not ev["is_open"]:
        return _closed(ev, 403)
    sent = request.form.get("csrf") or ""
    if not hmac.compare_digest(sent.encode(), _public_csrf(ev["slug"]).encode()):
        return _render_form(ev, request.form, notice="The form expired. Please submit it again.",
                            code=400)
    ip = _client_ip()
    if STORE.hits("ip:" + ip, IP_WINDOW) >= IP_LIMIT:
        return _render_form(ev, request.form, notice="Too many submissions from your network. "
                            "Please wait a few minutes and try again.", code=429)
    STORE.hit("ip:" + ip)
    if (request.form.get("website") or "").strip():
        # Honeypot filled in: a bot. Pretend it worked, store nothing.
        return redirect(url_for("invitation_public.thanks", slug=ev["slug"]), 303)
    value, errors = rules.parse_response(request.form)
    if errors:
        return _render_form(ev, request.form, errors, code=422)
    email_key = "em:%d:%s" % (ev["id"], hashlib.sha256(value["email_normalized"].encode()).hexdigest())
    if STORE.hits(email_key, EMAIL_WINDOW) >= EMAIL_LIMIT:
        return _render_form(ev, request.form, {"email": "This email was submitted too many times. "
                                               "Please try again in an hour."}, code=429)
    STORE.hit(email_key)
    STORE.upsert_response(ev["id"], value)
    return redirect(url_for("invitation_public.thanks", slug=ev["slug"]), 303)


@public_bp.get("/<slug>/thanks")
def thanks(slug):
    ev = _public_event(slug)
    return render_template("invitation/public_message.html", ev=ev, state="thanks")


@public_bp.get("/<slug>/image")
def image(slug):
    ev = _public_event(slug)
    path = _image_path(ev["image_key"])
    if path is None:
        abort(404)
    resp = send_file(path, mimetype=ev["image_mime"], conditional=True)
    resp.headers["Cache-Control"] = "public, max-age=86400"
    return resp
