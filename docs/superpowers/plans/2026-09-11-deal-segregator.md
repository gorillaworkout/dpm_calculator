# Deal Segregator Integration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the existing D&W Calculator Flask app into Dupoin DPM Tools with unchanged D&W behavior plus Deal Segregator producing one Excel workbook with Daily, Monthly Summary, and Verifikasi sheets.

**Architecture:** Keep one Flask/Gunicorn/PM2 service and extend its existing tool registry and asynchronous disk-backed job runner. Add the proven Deal Segregator engine as one focused module, then make upload validation, subprocess arguments, output names, and MIME types tool-specific with minimal conditionals.

**Tech Stack:** Python 3, Flask, Werkzeug, openpyxl, Gunicorn, PM2, HTML/Jinja, assert-based integration tests.

**Spec:** `docs/superpowers/specs/2026-09-11-deal-segregator-design.md`

## Global Constraints

- Production remains PM2 `dw-calculator`, Gunicorn on `127.0.0.1:3022`.
- No database, new service, port, domain, proxy, authentication, or saved history.
- Existing D&W calculations and inputs remain unchanged.
- Deal Segregator accepts one or more `.csv`, `.xlsx`, or `.xlsm` Deals History files.
- Deal Segregator returns one `.xlsx` containing `Daily`, `Monthly Summary`, `Verifikasi`.
- Country remains exactly as supplied; no lookup or enrichment.
- User-facing UI copy remains English.
- Disk-backed asynchronous state and six-hour cleanup remain in use.
- Financial source files, templates, generated workbooks, archives, caches, and virtual environments remain excluded from Git.

---

### Task 1: Establish Git Baseline and Safe Ignore Rules

**Files:**
- Create: `.gitignore`
- Modify: `README.md`
- Include: `docs/superpowers/specs/2026-09-11-deal-segregator-design.md`
- Include: `docs/superpowers/plans/2026-09-11-deal-segregator.md`

**Interfaces:**
- Consumes: Existing local source tree and empty `gorillaworkout/dpm_calculator` repository.
- Produces: `main` branch with a reproducible source-only baseline.

- [ ] **Step 1: Initialize Git and inspect ignored files**

Run:

```bash
git init
git branch -M main
git status --short --ignored
```

Expected: `.xlsx`, `.xlsm`, `.csv`, `.zip`, `.pdf`, caches, and app bundles show as ignored; source, templates, tests, docs, README, and `.gitignore` remain trackable.

- [ ] **Step 2: Add repository identity without replacing existing documentation**

Append this section to `README.md`:

```markdown
## Repository

GitHub: https://github.com/gorillaworkout/dpm_calculator
```

- [ ] **Step 3: Stage and audit the exact baseline**

Run:

```bash
git add .gitignore README.md BACA\ DULU.txt app.py buat_template.py gabung_fee.py hitung_dw.py isi_template.py mtoatd_spec.py rapikan_file.py tambah_rate.py test_app.py test_regressions.py templates verifikasi docs
git status --short
git diff --cached --stat
git diff --cached --check
```

Expected: no transaction fixture, generated workbook, ZIP, PDF, `.venv`, cache, or secret file staged.

- [ ] **Step 4: Commit and push the baseline**

Run:

```bash
git commit -m "chore: establish DPM calculator baseline"
git remote add origin https://github.com/gorillaworkout/dpm_calculator.git
git push -u origin main
```

Expected: GitHub `main` points to the local baseline commit.

---

### Task 2: Write the Failing Deal Segregator Engine Test

**Files:**
- Create: `test_deal_segregator.py`
- Create later: `deal_segregator.py`
- Fixture: `../zern/31 Jul_Deals History 2026_08_07 12_07_27 (1).csv` (local only, never committed)

**Interfaces:**
- Consumes: `proses(paths_in: list[Path], path_out: Path) -> dict`.
- Produces: Contract for one workbook containing exactly three required sheets and deduplicated aggregate data.

- [ ] **Step 1: Write a minimal failing engine test**

Create `test_deal_segregator.py` with assertions equivalent to:

```python
from pathlib import Path
from tempfile import TemporaryDirectory
from openpyxl import load_workbook
from deal_segregator import proses

SRC = Path("/Users/bayudarmawan/Documents/Dupoin/zern/31 Jul_Deals History 2026_08_07 12_07_27 (1).csv")

assert SRC.is_file()
with TemporaryDirectory() as tmp:
    out = Path(tmp) / "hasil.xlsx"
    summary = proses([SRC], out)
    wb = load_workbook(out, read_only=True, data_only=True)
    assert wb.sheetnames == ["Daily", "Monthly Summary", "Verifikasi"]
    assert wb["Daily"].max_row - 1 == 9418
    assert wb["Monthly Summary"].max_row - 1 == 9418
    assert summary["dipakai"] == 209838

with TemporaryDirectory() as tmp:
    out = Path(tmp) / "dedupe.xlsx"
    summary = proses([SRC, SRC], out)
    assert summary["dipakai"] == 209838

print("OK deal segregator engine")
```

- [ ] **Step 2: Run it and verify RED**

Run:

```bash
python3 test_deal_segregator.py
```

Expected: FAIL with `ModuleNotFoundError: No module named 'deal_segregator'`.

- [ ] **Step 3: Commit only the failing test**

```bash
git add test_deal_segregator.py
git commit -m "test: define deal segregator workbook contract"
```

---

### Task 3: Implement the One-Workbook Deal Segregator Engine

**Files:**
- Create: `deal_segregator.py`
- Test: `test_deal_segregator.py`
- Reference only: `/Users/bayudarmawan/Documents/Dupoin/zern/deal_segregator.py`

**Interfaces:**
- Consumes: Deals History paths and an `.xlsx` output path.
- Produces: `proses(paths_in, path_out)` summary plus one workbook with `Daily`, `Monthly Summary`, `Verifikasi`.

- [ ] **Step 1: Copy the validated parsing and aggregation logic**

Copy these source behaviors unchanged: encoding/separator detection, required-column validation, normalization, Deal ID deduplication, daily/monthly aggregation, sorting, styles, and verification metrics.

- [ ] **Step 2: Replace ZIP output with one workbook writer**

Use one writer with this exact shape:

```python
def _tulis_workbook(harian, bulanan, path_out: Path, ringkasan: dict):
    wb = Workbook()
    daily = wb.active
    daily.title = "Daily"
    _tulis_sheet(daily, harian, "Date", "dd mmm yyyy")
    monthly = wb.create_sheet("Monthly Summary")
    _tulis_sheet(monthly, bulanan, "Month", "mmm yyyy")
    _tulis_verifikasi(wb.create_sheet("Verifikasi"), ringkasan)
    path_out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path_out)
```

Call `_tulis_workbook(...)` from `proses`; set CLI default filename to `<input>-hasil.xlsx`; describe `.xlsx` rather than `.zip` in module and argparse help.

- [ ] **Step 3: Run engine tests and verify GREEN**

Run:

```bash
python3 test_deal_segregator.py
python3 -m py_compile deal_segregator.py
```

Expected: both pass; workbook has exactly three sheets in required order.

- [ ] **Step 4: Commit the engine**

```bash
git add deal_segregator.py
git commit -m "feat: generate segregated deals workbook"
```

---

### Task 4: Write Failing Flask Integration Tests

**Files:**
- Modify: `test_app.py`
- Modify later: `app.py`
- Modify later: `templates/base.html`
- Modify later: `templates/index.html`
- Modify later: `templates/tool.html`

**Interfaces:**
- Consumes: Flask routes `/`, `/tool/dw`, `/tool/segregate`, `/job/<id>/download`.
- Produces: Contract proving both tools coexist and use the correct inputs/output type.

- [ ] **Step 1: Replace the one-menu assertion**

Assert:

```python
assert set(TOOLS) == {"dw", "segregate"}
assert b"Dupoin DPM Tools" in c.get("/").data
assert b"Generate D&amp;W" in c.get("/").data
assert b"Deal Segregator" in c.get("/").data
```

- [ ] **Step 2: Add Deal Segregator form and validation assertions**

Assert:

```python
page = c.get("/tool/segregate")
assert page.status_code == 200
assert b'name="file"' in page.data
assert b"multiple" in page.data
assert b'name="bulan"' not in page.data
assert b'name="saldo_jw"' not in page.data
assert c.post("/tool/segregate", data={}, content_type="multipart/form-data").status_code == 400
assert c.post(
    "/tool/segregate",
    data={"file": [(io.BytesIO(b"x"), "a.txt")]},
    content_type="multipart/form-data",
).status_code == 400
```

- [ ] **Step 3: Add one-file and duplicate-file async pipeline checks**

Post the zern fixture once, poll state, download, load `r.data` directly with `openpyxl.load_workbook(io.BytesIO(r.data))`, and assert exact sheet order. Repeat with the same fixture twice and assert the log still reports `dipakai: 209838`.

- [ ] **Step 4: Run and verify RED**

Run:

```bash
python3 test_app.py
```

Expected: FAIL because `TOOLS` lacks `segregate` and landing identity remains D&W Calculator.

- [ ] **Step 5: Commit failing integration tests**

```bash
git add test_app.py
git commit -m "test: define two-tool Flask workflow"
```

---

### Task 5: Integrate Deal Segregator into Flask

**Files:**
- Modify: `app.py`
- Modify: `templates/base.html`
- Modify: `templates/index.html`
- Modify: `templates/tool.html`
- Test: `test_app.py`

**Interfaces:**
- Consumes: `deal_segregator.py` CLI: `python deal_segregator.py <inputs...> -o <output.xlsx>`.
- Produces: tool-specific forms, validation, async subprocess invocation, filenames, and Excel responses.

- [ ] **Step 1: Register `segregate` and its documentation**

Add `TOOLS["segregate"] = ("Deal Segregator", ..., ("deal_segregator.py",), True)`, a Deals History hint, and the approved preparation/process/output copy from the design and zern reference.

- [ ] **Step 2: Make GET form context tool-specific**

Only pass `_pilihan_bulan()` for `slug == "dw"`. Pass empty/default context for `segregate`, preventing D&W month and wallet fields from rendering.

- [ ] **Step 3: Make POST validation tool-specific**

For `dw`, preserve current single-file validation, period, and balance behavior byte-for-byte where possible. For `segregate`, use:

```python
uploads = [u for u in request.files.getlist("file") if u and u.filename]
if not uploads:
    return _gagal("Please choose at least one .csv or .xlsx file first.")
invalid = [u.filename for u in uploads if not u.filename.lower().endswith((".csv", ".xlsx", ".xlsm"))]
if invalid:
    return _gagal(f"These files are not .csv/.xlsx and were rejected: {', '.join(invalid)}")
```

Save colliding upload names safely, preserving every file. Set output to `hasil.xlsx` and download name ending `-hasil.xlsx`.

- [ ] **Step 4: Generalize `_process_job` minimally**

Allow the first Deal Segregator step to receive a list of source paths, while D&W remains a single chained source. Keep period/balance flags D&W-only. Preserve disk state, exception capture, cleanup, and async thread behavior.

- [ ] **Step 5: Return Excel MIME type for both tools**

Use `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`; no ZIP branch remains.

- [ ] **Step 6: Adapt shared templates**

Brand shell as `Dupoin DPM Tools`, use `DPM` mark, show both nav items, retain template link only when on the D&W flow. Render the file label, `accept`, `multiple`, hints, month controls, and wallet controls based on `slug`.

- [ ] **Step 7: Update landing page**

Use generic Dupoin DPM Tools intro and two cards generated by `TOOLS`. Replace D&W-only explanatory sections with a concise explanation of both workflows.

- [ ] **Step 8: Run integration and regression tests**

Run:

```bash
python3 test_app.py
python3 test_deal_segregator.py
python3 test_regressions.py
python3 -m py_compile app.py deal_segregator.py
```

Expected: all pass with no traceback or warning.

- [ ] **Step 9: Commit Flask integration**

```bash
git add app.py templates/base.html templates/index.html templates/tool.html
git commit -m "feat: add Deal Segregator menu"
```

---

### Task 6: Local Final Verification and Push

**Files:**
- Verify all tracked source.

**Interfaces:**
- Consumes: Completed local app.
- Produces: clean, tested `main` ready for hand-sync deployment.

- [ ] **Step 1: Run complete checks again from a clean command**

```bash
python3 test_app.py && python3 test_deal_segregator.py && python3 test_regressions.py
```

Expected: exit code 0.

- [ ] **Step 2: Inspect repository state and secret patterns**

```bash
git status --short
git diff --check
git grep -nEi '(password|api[_-]?key|secret|token)[[:space:]]*=' -- ':!docs/**' || true
git ls-files '*.xlsx' '*.xlsm' '*.csv' '*.zip' '*.pdf'
```

Expected: no uncommitted implementation files, no credentials, no tracked financial fixtures.

- [ ] **Step 3: Push tested main**

```bash
git push origin main
```

Expected: remote main matches local HEAD.

---

### Task 7: Backup and Hand-Sync Deploy

**Files:**
- Deploy only changed/new runtime files among: `app.py`, `deal_segregator.py`, `templates/base.html`, `templates/index.html`, `templates/tool.html`.
- Do not deploy docs, tests, fixtures, caches, or ignored files unless required for a remote smoke check.

**Interfaces:**
- Consumes: Tested local runtime files.
- Produces: Updated `/home/ubuntu/apps/dw-calculator` with rollback archive and MD5 parity.

- [ ] **Step 1: Record local and remote fingerprints**

```bash
for f in app.py deal_segregator.py templates/base.html templates/index.html templates/tool.html; do
  test -f "$f" && md5 -q "$f"
done
ssh aws-marketing 'cd ~/apps/dw-calculator && md5sum app.py deal_segregator.py templates/base.html templates/index.html templates/tool.html 2>/dev/null || true'
```

- [ ] **Step 2: Create mandatory VPS backup**

```bash
ssh aws-marketing 'mkdir -p ~/backups && tar czf ~/backups/dw-calculator-pre-deal-segregator-$(date +%Y%m%d-%H%M%S).tgz --exclude=.venv --exclude=__pycache__ -C ~/apps dw-calculator'
```

Expected: archive path printed/verified with `test -s`.

- [ ] **Step 3: Upload only changed files**

Use `scp` for the exact MD5 diff. Templates go to `/home/ubuntu/apps/dw-calculator/templates/`; Python files go to the app root.

- [ ] **Step 4: Verify MD5 parity and import**

```bash
ssh aws-marketing 'cd ~/apps/dw-calculator && .venv/bin/python3 -m py_compile app.py deal_segregator.py && md5sum app.py deal_segregator.py templates/base.html templates/index.html templates/tool.html'
```

Expected: every remote hash equals local hash.

- [ ] **Step 5: Restart and persist PM2**

```bash
ssh aws-marketing 'cd ~/apps/dw-calculator && rm -rf __pycache__ && pm2 restart dw-calculator && pm2 save'
```

---

### Task 8: Verify Production Round Trips

**Files:**
- Temporary local outputs only; never commit.

**Interfaces:**
- Consumes: Public `https://d&w-calculator.gorillaworkout.id` routes and fixture upload.
- Produces: Evidence that both routes and the Deal Segregator result work after restart.

- [ ] **Step 1: Verify process, port, logs, direct app identity**

```bash
ssh aws-marketing 'pm2 describe dw-calculator; ss -ltnp | grep :3022; curl -fsS http://127.0.0.1:3022/ | grep -F "Dupoin DPM Tools"; pm2 logs dw-calculator --lines 30 --nostream'
```

Expected: PM2 online, port bound, correct identity, no startup traceback.

- [ ] **Step 2: Verify public routes with the literal hostname safely URL-encoded/shell-quoted**

```bash
curl -fsS 'https://d%26w-calculator.gorillaworkout.id/'
curl -fsS 'https://d%26w-calculator.gorillaworkout.id/tool/dw'
curl -fsS 'https://d%26w-calculator.gorillaworkout.id/tool/segregate'
```

If the encoded hostname is rejected by the client, use the known literal production URL exactly as configured and quote it to protect `&`. Assert landing marker and each tool-specific input marker, not status alone.

- [ ] **Step 3: Run public Deal Segregator E2E**

POST the real MT5 fixture to `/tool/segregate`, follow the returned `/job/<id>`, poll until done, download to a temporary `.xlsx`, then verify:

```python
from openpyxl import load_workbook
wb = load_workbook("/tmp/deal-segregator-public.xlsx", read_only=True, data_only=True)
assert wb.sheetnames == ["Daily", "Monthly Summary", "Verifikasi"]
assert wb["Daily"].max_row - 1 == 9418
assert wb["Monthly Summary"].max_row - 1 == 9418
```

- [ ] **Step 4: Verify restart resilience**

Restart once more, then repeat landing and both route marker probes. Confirm PM2 remains online.

- [ ] **Step 5: Report evidence**

Report GitHub URL and commit, backup archive, exact local/remote hashes, test outputs, PM2 status, route markers, workbook sheet names, row counts, and any public D&W full-pipeline fixture limitation.
