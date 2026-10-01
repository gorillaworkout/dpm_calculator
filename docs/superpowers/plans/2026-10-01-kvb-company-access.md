# KVB Company Access and D&W Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add per-account DPM/KVB access and one isolated KVB Generate D&W workflow without changing existing DPM behavior.

**Architecture:** Preserve the existing DPM `TOOLS`, URLs, forms, subprocess list, outputs, and tests as a regression boundary. Add company access columns to the existing SQLite table, authorize company-qualified KVB routes server-side, and run KVB through a separate three-script definition that reuses the existing disk-backed job runner. Keep templates shared only where rendering differences are data-driven; do not alter calculator scripts.

**Tech Stack:** Python 3, Flask, SQLite, Jinja2, openpyxl, subprocess, assert-based Python test scripts.

**Spec:** `docs/superpowers/specs/2026-10-01-kvb-company-access-design.md`

## Global Constraints

- Never delete, recreate, reset, clean, or overwrite `users.db` or existing user work.
- Existing DPM URLs remain valid: `/`, `/tool/dw`, `/tool/segregate`, `/job/<id>`, `/job/<id>/download`, `/template`.
- Existing DPM `Generate D&W` pipeline remains exactly `isi_template.py -> hitung_dw.py`.
- Deal Segregator remains DPM-only; do not modify `deal_segregator.py`.
- Do not modify `isi_template.py`, `hitung_dw.py`, or `siapkan_kvb.py` calculation/translation behavior.
- KVB pipeline is exactly `siapkan_kvb.py -> isi_template.py -> hitung_dw.py`.
- User-facing UI copy remains English.
- Direct unauthorized company URLs return HTTP 403.
- Existing admins migrate to DPM+KVB; existing normal users migrate to DPM-only.
- New or edited accounts must retain at least one company.
- Current worktree is dirty. Stage only task-specific hunks/files. Never use `git reset`, `git clean`, `git checkout --`, `git restore`, or stash.
- Because `app.py`, `auth.py`, and templates already contain uncommitted user work, do not create checkpoint commits that would silently absorb unrelated changes. Record RED/GREEN commands in the TDD evidence file; make a final commit only after reviewing the exact staged diff with the user if requested.

## File Structure

- Modify `auth.py`: additive SQLite migration, company-access helpers, admin create/update validation.
- Modify `app.py`: separate KVB tool metadata/routes, server authorization, KVB job ownership metadata, unchanged DPM pipeline.
- Modify `templates/base.html`: company switcher and data-driven company branding/navigation.
- Modify `templates/index.html`: render company-specific landing copy while preserving DPM content.
- Modify `templates/tool.html`: company-aware form links/copy; preserve existing DPM controls.
- Modify `templates/job.html`: company-aware job/download/retry URLs.
- Modify `templates/admin.html`: create/edit company access controls.
- Create `test_company_access.py`: migration, authorization, navigation, and admin-access tests.
- Create `test_kvb_web.py`: exact KVB pipeline/order/period/failure tests with subprocess isolation.
- Create `docs/testing/kvb-company-access.tdd.md`: actual RED/GREEN and regression evidence.

---

### Task 1: Migrate User Access Without Recreating the Database

**Files:**
- Modify: `auth.py:59-87, 230-288`
- Create: `test_company_access.py`

**Interfaces:**
- Produces: `COMPANY_COLUMNS = {"dpm": "can_dpm", "kvb": "can_kvb"}`.
- Produces: `has_company(user: sqlite3.Row, company: str) -> bool`.
- Produces route: `POST /admin/users/<username>/access` with form keys `can_dpm`, `can_kvb`.
- Existing `current_user()`, `verify()`, login cookies, password reset, and deletion semantics remain unchanged.

- [ ] **Step 1: Write migration and admin validation tests**

Create `test_company_access.py`. Before importing `app`, create a legacy SQLite database containing the exact old schema and two hashed users:

```python
legacy.execute("""CREATE TABLE users (
    username TEXT PRIMARY KEY,
    pw_hash TEXT NOT NULL,
    is_admin INTEGER NOT NULL DEFAULT 0,
    created REAL NOT NULL,
    last_seen REAL)""")
legacy.execute("INSERT INTO users VALUES (?,?,?,?,?)",
               ("legacy-admin", generate_password_hash("admin-password-1"), 1, time.time(), None))
legacy.execute("INSERT INTO users VALUES (?,?,?,?,?)",
               ("legacy-user", generate_password_hash("user-password-1"), 0, time.time(), None))
```

After importing the app, assert:

```python
assert auth.verify("legacy-admin", "admin-password-1")["can_dpm"] == 1
assert auth.verify("legacy-admin", "admin-password-1")["can_kvb"] == 1
assert auth.verify("legacy-user", "user-password-1")["can_dpm"] == 1
assert auth.verify("legacy-user", "user-password-1")["can_kvb"] == 0
```

Also assert the legacy rows and hashes still exist, bootstrap admins receive both companies, account creation rejects no company, DPM-only/KVB-only/dual creation succeeds, access update changes only access columns, and password hashes remain byte-for-byte unchanged after access edits.

- [ ] **Step 2: Run the test and verify RED**

Run:

```bash
python3 test_company_access.py
```

Expected: FAIL because `can_dpm`, `can_kvb`, `has_company`, and `/admin/users/<username>/access` do not exist.

- [ ] **Step 3: Add idempotent SQLite migration**

In `init_db()`, retain `CREATE TABLE IF NOT EXISTS`, then inspect columns:

```python
columns = {row["name"] for row in con.execute("PRAGMA table_info(users)")}
if "can_dpm" not in columns:
    con.execute("ALTER TABLE users ADD COLUMN can_dpm INTEGER NOT NULL DEFAULT 1")
if "can_kvb" not in columns:
    con.execute("ALTER TABLE users ADD COLUMN can_kvb INTEGER NOT NULL DEFAULT 0")
con.execute("UPDATE users SET can_dpm=1, can_kvb=1 WHERE is_admin=1")
```

Extend the fresh-table schema with both non-null columns. Insert bootstrap admins with both values set to `1`. This migration is additive and idempotent; never issue `DROP TABLE` or delete rows.

- [ ] **Step 4: Add company helper and admin writes**

Add:

```python
COMPANY_COLUMNS = {"dpm": "can_dpm", "kvb": "can_kvb"}

def has_company(user, company):
    column = COMPANY_COLUMNS.get(company)
    return bool(user and column and user[column])
```

In `admin_create`, parse checkboxes and reject neither selected:

```python
can_dpm = 1 if request.form.get("can_dpm") else 0
can_kvb = 1 if request.form.get("can_kvb") else 0
if not (can_dpm or can_kvb):
    return _admin_error("Choose at least one company: DPM or KVB.")
```

Insert both access values. Add `POST /admin/users/<username>/access`; validate at least one, verify user exists, then update only `can_dpm` and `can_kvb`.

- [ ] **Step 5: Run the focused test and verify GREEN**

Run:

```bash
python3 test_company_access.py
```

Expected: `OK company access`.

- [ ] **Step 6: Run existing auth regressions**

Run:

```bash
python3 test_auth.py && python3 test_login.py
```

Expected: `OK auth` and `OK login`. If old create-user fixtures omit access fields, update test fixtures—not production defaults—to explicitly send `can_dpm=1`, preserving the production rule that new accounts choose access.

---

### Task 2: Add Server-Side Company Authorization Without Changing DPM Routes

**Files:**
- Modify: `app.py:64-110, 326-367, 530-595, 633-636`
- Extend: `test_company_access.py`

**Interfaces:**
- Consumes: `auth.has_company(user, company)`.
- Produces: `KVB_TOOLS` containing only `dw` with scripts `("siapkan_kvb.py", "isi_template.py", "hitung_dw.py")`.
- Produces: `company_required(company: str)` internal authorization helper.
- Produces routes: `GET/POST /kvb/tool/dw`, `GET /kvb/job/<job_id>`, `GET /kvb/job/<job_id>/download`.
- Preserves existing DPM route endpoint names and paths.

- [ ] **Step 1: Add failing authorization tests**

Extend `test_company_access.py` with authenticated clients for DPM-only, KVB-only, and dual users. Assert:

```python
assert dpm.get("/tool/dw").status_code == 200
assert dpm.get("/kvb/tool/dw").status_code == 403
assert kvb.get("/kvb/tool/dw").status_code == 200
assert kvb.get("/tool/dw").status_code == 403
assert both.get("/tool/segregate").status_code == 200
assert both.get("/kvb/tool/dw").status_code == 200
assert kvb.get("/tool/segregate").status_code == 403
```

Create job-state fixtures with `company="dpm"` and `company="kvb"`; assert status/download routes reject cross-company users before reading or delivering result files.

- [ ] **Step 2: Run focused test and verify RED**

Run:

```bash
python3 test_company_access.py
```

Expected: FAIL because KVB routes and DPM authorization do not exist.

- [ ] **Step 3: Define a separate KVB tool registry**

Keep current `TOOLS` unchanged. Add:

```python
KVB_TOOLS = {
    "dw": (
        "Generate D&W",
        "Upload a KVB Plus workbook and generate the finished monthly D&W report.",
        ("siapkan_kvb.py", "isi_template.py", "hitung_dw.py"),
        True,
    ),
}
```

Add company lookup functions rather than changing DPM tuple structure:

```python
def _tools_for(company):
    return TOOLS if company == "dpm" else KVB_TOOLS if company == "kvb" else None

def _require_company(company):
    if not auth.has_company(request.user, company):
        abort(403, description=f"Your account does not have access to {company.upper()}.")
```

- [ ] **Step 4: Gate existing DPM tool routes and add KVB routes**

After global login sets `request.user`, call `_require_company("dpm")` inside existing `/`, `/tool/<slug>`, tool POST, and `/template` handlers. Do not rename or redirect them.

Add `/kvb`, `/kvb/tool/<slug>` GET/POST wrappers or shared internal functions receiving `company`; route only through `KVB_TOOLS`. Reject unknown KVB slugs with 404. Ensure KVB cannot route to `segregate`.

- [ ] **Step 5: Bind jobs to owner company**

Persist `company` in every state record. Existing DPM jobs use `company="dpm"`; KVB jobs use `company="kvb"`. Status and download handlers authorize against `info["company"]`, treating old state files without the key as DPM for compatibility:

```python
company = info.get("company", "dpm")
_require_company(company)
```

Add company-qualified KVB job status/download routes while preserving old DPM routes. Generate links with endpoint names rather than hardcoded `/job/...` paths.

- [ ] **Step 6: Run focused and DPM route tests**

Run:

```bash
python3 test_company_access.py && python3 test_login.py
```

Expected: both pass. Verify existing login `next=/tool/dw` remains valid.

---

### Task 3: Render Company Switcher and Admin Access Controls

**Files:**
- Modify: `templates/base.html:1-214`
- Modify: `templates/index.html:1-87`
- Modify: `templates/tool.html:1-410`
- Modify: `templates/job.html:1-110`
- Modify: `templates/admin.html:1-81`
- Extend: `test_company_access.py`

**Interfaces:**
- Consumes template values: `company`, `company_tools`, `company_urls`, and `request.user.can_dpm/can_kvb`.
- Existing DPM template context defaults to `company="dpm"` and current `TOOLS`.
- KVB links use Flask-generated URLs supplied by views, never string replacement of DPM URLs.

- [ ] **Step 1: Add failing rendering tests**

Assert dual-access sidebar contains enabled links for DPM and KVB; DPM-only contains an enabled DPM control plus a disabled KVB control without an `href`; KVB-only has the inverse. Assert KVB page includes `KVB`, `KVB Plus`, and only `Generate D&W`; it must not contain `Deal Segregator`.

Assert DPM snapshots retain:

```python
assert b"Dupoin DPM Tools" in dpm_page.data
assert b"Generate D&amp;W" in dpm_page.data
assert b"Deal Segregator" in dpm_page.data
assert b'href="/tool/dw"' in dpm_page.data
assert b'href="/tool/segregate"' in dpm_page.data
```

Assert admin create form has `can_dpm` and `can_kvb`, and each user row has an access-edit form.

- [ ] **Step 2: Run focused test and verify RED**

Run:

```bash
python3 test_company_access.py
```

Expected: FAIL on missing switcher/admin fields/KVB copy.

- [ ] **Step 3: Add company-aware template context**

Change `_nav()` to return current company tools and company switch metadata while retaining `tools` for current template loops. Default non-company admin/login contexts safely to the first allowed company.

The switcher renders both companies. Enabled examples:

```html
<a href="{{ company_urls.dpm }}" aria-current="{{ 'page' if company == 'dpm' }}">DPM</a>
```

Disabled examples:

```html
<span aria-disabled="true" title="Your account does not have KVB access">KVB</span>
```

Do not use client-side checks for authorization.

- [ ] **Step 4: Preserve DPM copy; add KVB branches only**

In `base.html`, keep DPM brand text for `company == "dpm"`; render `KVB Tools`/`KVB` only for KVB. Keep the existing DPM nav hrefs unchanged. In `index.html`, wrap current content unchanged in the DPM branch and add a concise KVB landing branch describing KVB Plus upload and the three internal stages.

In `tool.html`, preserve existing DPM branches and fields. For KVB `dw`, change only company-specific labels/hints and submit target supplied by the view. Keep report month and supported opening-balance fields because downstream scripts are unchanged.

In `job.html`, replace hardcoded DPM endpoint/path references with supplied `job_url`, `download_url`, and `new_job_url`, yielding the same old DPM URLs and KVB-qualified URLs.

- [ ] **Step 5: Add admin access inputs**

In account creation, add checked DPM and optional KVB checkboxes. In each user row, show DPM/KVB badges and an inline access form posting to `admin_access`. Preserve password reset/delete forms.

- [ ] **Step 6: Run rendering and auth regressions**

Run:

```bash
python3 test_company_access.py && python3 test_auth.py && python3 test_login.py
```

Expected: all pass.

---

### Task 4: Run the Isolated KVB Three-Stage Pipeline

**Files:**
- Modify: `app.py:113-222, 344-527`
- Create: `test_kvb_web.py`

**Interfaces:**
- Consumes: `KVB_TOOLS["dw"][2]` exact ordered tuple.
- Produces KVB final download name: `<original stem> - <Mon YYYY>-hasil.xlsx` with KVB company context in state.
- Preserves `_process_job()` support for DPM and Deal Segregator.

- [ ] **Step 1: Write exact-order pipeline tests**

Create `test_kvb_web.py` with a temporary auth DB, KVB-capable user, job directory, and patched `subprocess.run`. The fake runner records each command and creates the `-o` file. POST a small fake `.xlsx` to `/kvb/tool/dw` with August 2026.

Assert commands are exactly ordered:

```python
assert [Path(cmd[1]).name for cmd in commands] == [
    "siapkan_kvb.py", "isi_template.py", "hitung_dw.py"
]
assert "--bulan" not in commands[0]
assert commands[1][-2:] == ["--bulan", "2026-08"]
assert "--period" in commands[2]
assert commands[2][commands[2].index("--period") + 1] == "2026-08"
assert Path(commands[1][2]) == Path(commands[0][commands[0].index("-o") + 1])
assert Path(commands[2][2]) == Path(commands[1][commands[1].index("-o") + 1])
```

Also assert DPM POST still records only `isi_template.py`, `hitung_dw.py`.

- [ ] **Step 2: Add failure-short-circuit test**

Make fake `siapkan_kvb.py` return non-zero with `stderr="STOP: sheet Xero not found"`. Assert only one command ran, state is `failed`, HTTP status page is 422, and the message is visible. Make the second stage fail in another case; assert `hitung_dw.py` never runs.

- [ ] **Step 3: Run KVB tests and verify RED**

Run:

```bash
python3 test_kvb_web.py
```

Expected: FAIL because the KVB POST is not wired to job processing.

- [ ] **Step 4: Generalize request handling by explicit company**

Extract the shared current GET/POST mechanics into internal functions such as:

```python
def _render_tool(company, slug): ...
def _run_tool(company, slug): ...
```

These receive the registry from `_tools_for(company)`. Existing DPM route handlers call them with `"dpm"`; KVB handlers call them with `"kvb"`. Never infer company from the slug.

Pass `company` into the background thread and `_process_job(job_id, company, slug, ...)`. Select `steps = _tools_for(company)[slug][2]`. This explicit key prevents KVB scripts from entering the DPM path.

- [ ] **Step 5: Keep period/opening argument mapping script-based**

Retain:

```python
PERIODE_ARG = {"isi_template.py": "--bulan", "hitung_dw.py": "--period"}
SALDO_ARG = {"hitung_dw.py": "--jwallet-opening"}
CHANNEL_ARG = {"hitung_dw.py": "--channel-opening"}
```

Because `siapkan_kvb.py` is absent from these maps, it receives no unsupported period/balance flags. Intermediate names remain unique (`step0.xlsx`, `step1.xlsx`); final output is `dst`.

- [ ] **Step 6: Run focused KVB test and verify GREEN**

Run:

```bash
python3 test_kvb_web.py
```

Expected: `OK KVB web pipeline`.

- [ ] **Step 7: Run real translator unit checks**

Run the existing/new direct translator self-checks if present:

```bash
python3 -m py_compile siapkan_kvb.py app.py auth.py
python3 - <<'PY'
from siapkan_kvb import normalisasi_source_name
assert normalisasi_source_name("07-2026 Withdrawal") == "Withdrawal"
assert normalisasi_source_name("08-2026 Wallet withdrawal") == "Rebate Withdrawal"
assert normalisasi_source_name("new variant") == "new variant"
print("OK KVB translator mapping")
PY
```

Expected: compile succeeds and `OK KVB translator mapping`.

---

### Task 5: Prove DPM Non-Regression and Record Evidence

**Files:**
- Modify only if needed for explicit access fixture setup: `test_app.py`, `test_auth.py`, `test_login.py`, `test_chunked.py`, `test_chunked_e2e.py`, `test_regressions.py`
- Create: `docs/testing/kvb-company-access.tdd.md`

**Interfaces:**
- Consumes all prior behavior.
- Produces auditable RED/GREEN/regression evidence.

- [ ] **Step 1: Capture exact DPM invariants**

Before broad tests, run an introspection check:

```bash
python3 - <<'PY'
import app
assert app.TOOLS["dw"][2] == ("isi_template.py", "hitung_dw.py")
assert app.TOOLS["segregate"][2] == ("deal_segregator.py",)
assert set(app.TOOLS) == {"dw", "segregate"}
print("OK DPM registry unchanged")
PY
```

Expected: `OK DPM registry unchanged`.

- [ ] **Step 2: Run all lightweight auth/web/unit regressions**

Run:

```bash
python3 test_company_access.py
python3 test_kvb_web.py
python3 test_auth.py
python3 test_login.py
python3 test_chunked.py
python3 test_regressions.py
python3 test_deal_segregator.py
python3 test_deal_equity_fx.py
```

Expected: each script exits 0. Do not report PASS for missing fixture/dependency failures; record the exact blocker.

- [ ] **Step 3: Run integration tests with available fixtures**

Run:

```bash
python3 test_app.py
python3 test_chunked_e2e.py
```

Expected: exit 0 when referenced local fixtures are available. Confirm the DPM real workbook still generates an `.xlsx`, and Deal Segregator still generates the same ZIP members/count assertions.

- [ ] **Step 4: Verify no calculator or Deal Segregator code changed for this feature**

Record baseline paths and inspect the feature diff:

```bash
git diff --name-only 0a9b7db --
git diff -- hitung_dw.py isi_template.py deal_segregator.py siapkan_kvb.py
```

Expected for feature work: no new KVB-feature diff in `hitung_dw.py`, `isi_template.py`, `deal_segregator.py`, or `siapkan_kvb.py`. Existing pre-feature dirty diffs may appear; compare against the recorded pre-task status and do not claim they were introduced here.

- [ ] **Step 5: Write TDD evidence report**

Create `docs/testing/kvb-company-access.tdd.md` containing:

- Source spec and plan links.
- User journeys: DPM-only, KVB-only, dual-access admin, KVB upload.
- RED command/output excerpts for Tasks 1–4.
- GREEN command/output excerpts.
- DPM regression commands/results.
- Table mapping each guarantee to its test file and command.
- Known gaps, especially unavailable real KVB workbook fixtures.
- Statement that DB migration is additive and no DB was deleted.

Use only actual command output; never synthesize test results.

- [ ] **Step 6: Review final diff and working tree**

Run:

```bash
git diff --stat
git diff -- auth.py app.py templates/base.html templates/index.html templates/tool.html templates/job.html templates/admin.html test_company_access.py test_kvb_web.py docs/testing/kvb-company-access.tdd.md
git status --short --branch
```

Expected: changes are limited to access, company navigation, KVB web orchestration, tests, and evidence. No destructive DB operation. DPM registry and pipeline assertions remain green.

- [ ] **Step 7: Self-evaluate before delivery**

Apply `agent-self-evaluation`: score accuracy, completeness, clarity, actionability, and conciseness using test/diff evidence. Fix any score of 3 or below before reporting completion.
