#!/usr/bin/env python3
"""HTTP Basic Auth with a SQLite user store and an admin page.

Deliberately minimal: the browser's own Basic Auth dialog is the login form, so
there are no sessions, no cookies, no CSRF tokens and no login template. Users
are created and deleted from /admin by an admin account.

The very first admin is created from the environment on startup:
    DW_ADMIN_USER=admin DW_ADMIN_PASS=<secret>
Without those, no account exists and every request is denied -- which is the
safe failure mode.
"""
import os
import secrets
import sqlite3
import time
from functools import wraps
from pathlib import Path

from flask import (Response, current_app, redirect, render_template, request,
                   url_for)
from itsdangerous import (BadSignature, SignatureExpired, URLSafeTimedSerializer)
from werkzeug.security import check_password_hash, generate_password_hash

DB = Path(os.environ.get("DW_AUTH_DB", Path(__file__).resolve().parent / "users.db"))
COMPANY_COLUMNS = {"dpm": "can_dpm", "kvb": "can_kvb"}
REALM = 'Basic realm="Dupoin DPM Tools", charset="UTF-8"'
COOKIE = "dw_session"
SESSION_MAX_AGE = 12 * 3600          # a working day; then sign in again
SECRET_FILE = DB.with_name("session.key")


def secret_key():
    """A stable signing key, generated once and kept next to the user database.

    It must survive restarts, or every deploy would sign everybody out, and it
    must not live in the source tree.
    """
    try:
        return SECRET_FILE.read_text().strip()
    except OSError:
        SECRET_FILE.parent.mkdir(parents=True, exist_ok=True)
        key = secrets.token_hex(32)
        candidate = SECRET_FILE.with_name(f".{SECRET_FILE.name}.{os.getpid()}.{secrets.token_hex(8)}")
        try:
            fd = os.open(candidate, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, "w") as file:
                file.write(key)
                file.flush()
                os.fsync(file.fileno())
            try:
                os.link(candidate, SECRET_FILE)  # publish only after the full key exists
            except FileExistsError:
                pass
        finally:
            candidate.unlink(missing_ok=True)
        return SECRET_FILE.read_text().strip()


def _db():
    con = sqlite3.connect(DB, timeout=10)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with _db() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            pw_hash  TEXT NOT NULL,
            is_admin INTEGER NOT NULL DEFAULT 0,
            can_dpm INTEGER NOT NULL DEFAULT 1,
            can_kvb INTEGER NOT NULL DEFAULT 0,
            created  REAL NOT NULL,
            last_seen REAL)""")
        columns = {row["name"] for row in con.execute("PRAGMA table_info(users)")}
        if "can_dpm" not in columns:
            con.execute("ALTER TABLE users ADD COLUMN can_dpm INTEGER NOT NULL DEFAULT 1")
        if "can_kvb" not in columns:
            con.execute("ALTER TABLE users ADD COLUMN can_kvb INTEGER NOT NULL DEFAULT 0")
            # One-time legacy migration. Later admin access edits must survive restarts.
            con.execute("UPDATE users SET can_dpm=1, can_kvb=1 WHERE is_admin=1")
    # Bootstrap the first admin from env vars, once.
    user = os.environ.get("DW_ADMIN_USER")
    pw = os.environ.get("DW_ADMIN_PASS")
    if user and pw:
        with _db() as con:
            exists = con.execute("SELECT 1 FROM users WHERE username=?", (user,)).fetchone()
            if not exists:
                con.execute("INSERT INTO users (username, pw_hash, is_admin, can_dpm, can_kvb, created)"
                            " VALUES (?,?,1,1,1,?)",
                            (user, generate_password_hash(pw), time.time()))


def _find(username):
    with _db() as con:
        return con.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()


def has_company(user, company):
    column = COMPANY_COLUMNS.get(company)
    return bool(user and column and user[column])


def verify(username, password):
    """Return the user row when the credentials match, else None."""
    row = _find(username or "")
    # check_password_hash on a dummy hash keeps the timing similar for unknown
    # usernames, so the response time does not reveal which names exist.
    stored = row["pw_hash"] if row else "pbkdf2:sha256:600000$x$" + "0" * 64
    ok = check_password_hash(stored, password or "")
    if not (row and ok):
        return None
    with _db() as con:
        con.execute("UPDATE users SET last_seen=? WHERE username=?", (time.time(), username))
    return row


def _deny():
    return Response("Authentication required.", 401, {"WWW-Authenticate": REALM})


# --------------------------------------------------------------------- session
# The browser's own Basic Auth dialog cannot be styled and cannot be signed out
# of, so a real login page backed by a signed cookie is used instead. Basic Auth
# still works for scripts and tests -- both paths end at verify().

def current_user():
    """The signed-in user for this request, or None.

    A signed cookie is checked first, then an Authorization header. The cookie
    carries only the username; the account is re-read from the database on every
    request, so deleting a user or changing a password takes effect immediately
    instead of waiting for the cookie to expire.
    """
    if getattr(request, "_dw_user", None) is not None:
        return request._dw_user or None

    user = None
    token = request.cookies.get(COOKIE)
    if token:
        try:
            username = _signer().loads(token, max_age=SESSION_MAX_AGE)
            user = _find(username)
        except (BadSignature, SignatureExpired):
            user = None
    if user is None:
        head = request.authorization
        if head:
            user = verify(head.username, head.password)

    request._dw_user = user or False
    return user


def _signer():
    return URLSafeTimedSerializer(current_app.secret_key, salt="dw-login")


def issue_session(response, username):
    response.set_cookie(
        COOKIE, _signer().dumps(username),
        max_age=SESSION_MAX_AGE,
        httponly=True,           # JavaScript cannot read it, so XSS cannot steal it
        secure=True,             # the site is HTTPS-only behind Cloudflare
        samesite="Lax",          # not sent on cross-site POSTs, which blocks CSRF
    )
    return response


def clear_session(response):
    response.delete_cookie(COOKIE, samesite="Lax", secure=True, httponly=True)
    return response


def _safe_next(target):
    """Where to land after signing in.

    ?next= arrives from the URL, so an attacker could point it at their own site
    and use our login page as a springboard. Only same-site paths are honoured.
    """
    if not target or not target.startswith("/") or target.startswith("//"):
        return "/"
    if target.startswith("/login") or target.startswith("/logout"):
        return "/"
    return target


def _landing(user, target):
    target = _safe_next(target)
    if target == "/" and not has_company(user, "dpm") and has_company(user, "kvb"):
        return "/kvb"
    return target


def login_required(view):
    @wraps(view)
    def wrapped(*a, **kw):
        user = current_user()
        if not user:
            return _deny()
        request.user = user
        return view(*a, **kw)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*a, **kw):
        user = current_user()
        if not user:
            return redirect(url_for("login", next=request.path))
        if not user["is_admin"]:
            return render_template("login.html",
                                   error="That account cannot manage users."), 403
        request.user = user
        return view(*a, **kw)
    return wrapped


def register(app):
    """Attach the login, logout and admin routes to an existing Flask app."""
    init_db()
    app.secret_key = secret_key()

    @app.get("/login")
    def login():
        user = current_user()
        if user:
            return redirect(_landing(user, request.args.get("next")))
        return render_template("login.html", next=request.args.get("next", ""))

    @app.post("/login")
    def login_post():
        username = (request.form.get("username") or "").strip()
        user = verify(username, request.form.get("password") or "")
        if not user:
            # One message for both wrong-user and wrong-password on purpose:
            # saying which was wrong tells an attacker that a name is valid.
            time.sleep(0.4)          # blunt the speed of an automated guesser
            return render_template("login.html",
                                   error="Wrong username or password.",
                                   username=username,
                                   next=request.form.get("next", "")), 401
        target = _landing(user, request.form.get("next"))
        return issue_session(redirect(target), user["username"])

    @app.get("/logout")
    @app.post("/logout")
    def logout():
        return clear_session(redirect(url_for("login")))

    @app.get("/admin")
    @admin_required
    def admin_page():
        with _db() as con:
            users = con.execute("SELECT * FROM users ORDER BY is_admin DESC, username").fetchall()
        return render_template("admin.html", users=users, me=request.user["username"])

    @app.post("/admin/users")
    @admin_required
    def admin_create():
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        is_admin = 1 if request.form.get("is_admin") else 0
        can_dpm = 1 if request.form.get("can_dpm") else 0
        can_kvb = 1 if request.form.get("can_kvb") else 0
        if not username or not password:
            return _admin_error("Username and password are both required.")
        if len(password) < 10:
            return _admin_error("The password must be at least 10 characters long.")
        if not (can_dpm or can_kvb):
            return _admin_error("Choose at least one company: DPM or KVB.")
        if not username.replace("_", "").replace("-", "").replace(".", "").isalnum():
            return _admin_error("The username may only contain letters, digits, . - and _")
        with _db() as con:
            if con.execute("SELECT 1 FROM users WHERE username=?", (username,)).fetchone():
                return _admin_error(f"The user '{username}' already exists.")
            con.execute("INSERT INTO users (username, pw_hash, is_admin, can_dpm, can_kvb, created)"
                        " VALUES (?,?,?,?,?,?)",
                        (username, generate_password_hash(password), is_admin,
                         can_dpm, can_kvb, time.time()))
        return redirect(url_for("admin_page"))

    @app.post("/admin/users/<username>/access")
    @admin_required
    def admin_access(username):
        can_dpm = 1 if request.form.get("can_dpm") else 0
        can_kvb = 1 if request.form.get("can_kvb") else 0
        if not (can_dpm or can_kvb):
            return _admin_error("Choose at least one company: DPM or KVB.")
        with _db() as con:
            if not con.execute("SELECT 1 FROM users WHERE username=?", (username,)).fetchone():
                return _admin_error(f"No user named '{username}'.")
            con.execute("UPDATE users SET can_dpm=?, can_kvb=? WHERE username=?",
                        (can_dpm, can_kvb, username))
        return redirect(url_for("admin_page"))

    @app.post("/admin/users/<username>/delete")
    @admin_required
    def admin_delete(username):
        if username == request.user["username"]:
            return _admin_error("You cannot delete the account you are signed in with.")
        with _db() as con:
            admins = con.execute("SELECT COUNT(*) c FROM users WHERE is_admin=1").fetchone()["c"]
            target = con.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
            if not target:
                return _admin_error(f"No user named '{username}'.")
            if target["is_admin"] and admins <= 1:
                return _admin_error("This is the last admin account; create another one first.")
            con.execute("DELETE FROM users WHERE username=?", (username,))
        return redirect(url_for("admin_page"))

    @app.post("/admin/users/<username>/password")
    @admin_required
    def admin_reset(username):
        password = request.form.get("password") or ""
        if len(password) < 10:
            return _admin_error("The password must be at least 10 characters long.")
        with _db() as con:
            if not con.execute("SELECT 1 FROM users WHERE username=?", (username,)).fetchone():
                return _admin_error(f"No user named '{username}'.")
            con.execute("UPDATE users SET pw_hash=? WHERE username=?",
                        (generate_password_hash(password), username))
        return redirect(url_for("admin_page"))

    def _admin_error(message):
        with _db() as con:
            users = con.execute("SELECT * FROM users ORDER BY is_admin DESC, username").fetchall()
        return render_template("admin.html", users=users, me=request.user["username"],
                               error=message), 400
