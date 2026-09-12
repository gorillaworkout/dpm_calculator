# Task 4–5 Review Fix Report

## Scope

Review fixes only. D&W calculation code remains unchanged. No deployment or push performed. The financial Deals History fixture remains external and untracked:

`/Users/bayudarmawan/Documents/Dupoin/zern/31 Jul_Deals History 2026_08_07 12_07_27 (1).csv`

## Changes

### Real-fixture asynchronous E2E

`test_app.py` now uploads the external 121,959,192-byte fixture through `/tool/segregate`, polls the disk-backed asynchronous job state, downloads the generated workbook, and asserts:

- exact sheet order: `Daily`, `Monthly Summary`, `Verifikasi`;
- `Daily` has exactly 9,418 data rows;
- `Monthly Summary` has exactly 9,418 data rows.

The same fixture is then uploaded twice. The test asserts the subprocess log contains exactly the required count fragment `dipakai: 209838`, proving duplicate Deal IDs are not double-counted.

The real duplicate multipart request is approximately 244 MB, so the pre-existing 200 MB Flask request ceiling returned HTTP 413. `MAX_CONTENT_LENGTH` was minimally raised to 500 MB to allow the required supported multi-file workflow while retaining a finite upload limit.

### Unicode filename suffix regression

Root cause: Werkzeug `secure_filename("交易.xlsx")` returns `"xlsx"`. That sanitized value has no `.xlsx` suffix, so the downstream reader cannot identify the already-validated workbook type.

The upload path now appends the lower-cased, previously validated original suffix only when the sanitized filename has no suffix. Existing sanitized names and collision numbering remain unchanged.

## Strict TDD Evidence

### RED

Regression test added first. It posts an in-memory upload named `交易.xlsx`, intercepts background-thread construction before processing, and inspects the actual saved source path passed to `_process_job`.

Command:

```text
python3 test_app.py
```

Expected failure observed, exit 1:

```text
AssertionError: /private/var/folders/.../dw-calculator-jobs/37d5d1592c4749b08806ba352956ccf3/xlsx
```

The failure was the targeted behavior: saved path suffix was empty instead of `.xlsx`.

### GREEN

Minimal production change:

```python
safe = secure_filename(upload.filename) or f"upload-{index}"
if not Path(safe).suffix:
    safe += Path(upload.filename).suffix.lower()
```

The focused regression then passed as part of `python3 test_app.py`. During the full real-fixture run, the duplicate upload initially exposed HTTP 413 from the old 200 MB limit. After raising the finite ceiling to 500 MB, the entire app test passed.

Final GREEN output:

```text
OK 2 menu, 718,394 bytes hasil Jun 2026
```

## Required Verification

All requested commands passed from the feature worktree:

```text
python3 test_app.py
  exit 0 — OK 2 menu, 718,394 bytes hasil Jun 2026

python3 test_deal_segregator.py
  exit 0 — OK deal segregator engine

python3 test_regressions.py
  exit 0 — no output

python3 -m py_compile app.py deal_segregator.py
  exit 0 — no output
```

Additional check:

```text
git diff --check
  exit 0
```

## Concerns

- `test_app.py` intentionally depends on the absolute external financial fixture path; it will fail clearly when that untracked fixture is unavailable.
- The duplicate E2E transmits roughly 244 MB and performs two full aggregations, so it is resource-intensive.
- The 500 MB Flask ceiling permits the required duplicate fixture while preserving a bounded request size.
- No D&W calculation files changed. No push or deployment performed.

## Important Security Fix Round

### Change

Restored the unauthenticated production request ceiling in `app.py` to `200 * 1024 * 1024`. The required duplicate real-fixture HTTP test now raises `app.config["MAX_CONTENT_LENGTH"]` to `500 * 1024 * 1024` in `test_app.py` before creating the test client. Production behavior remains bounded at 200 MB; only the test process accepts the approximately 244 MB multipart request.

### Strict TDD Evidence

RED added first:

```python
assert app.config["MAX_CONTENT_LENGTH"] == 200 * 1024 * 1024
```

Command and targeted failure:

```text
python3 test_app.py
  exit 1 — AssertionError at test_app.py:17 because production still configured 500 MB
```

Minimal production fix changed only the default ceiling. GREEN:

```text
python3 test_app.py
  exit 0 — OK 2 menu, 718,394 bytes hasil Jun 2026
```

The GREEN run includes the mandatory duplicate external-fixture HTTP upload and `dipakai: 209838` assertion.

### Fix-round Verification

```text
python3 test_deal_segregator.py
  exit 0 — OK deal segregator engine

python3 test_regressions.py
  exit 0 — no output

python3 -m py_compile app.py deal_segregator.py test_app.py
  exit 0 — no output

git diff --check
  exit 0 — no output
```

### Fix-round Concerns

- `test_app.py` still requires the absolute external financial fixture and is intentionally resource-intensive.
- The test-only 500 MB override must remain before `app.test_client()` creation and must not move into production configuration.
- No D&W calculation files changed. No push or deployment performed.
