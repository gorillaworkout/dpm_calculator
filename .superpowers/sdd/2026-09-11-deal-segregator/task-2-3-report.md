# Tasks 2–3 Report

## Status

Complete. Deal Segregator engine added without Flask changes.

## TDD evidence

### RED

Command: `python3 test_deal_segregator.py`

Result: exit 1, expected failure:

```text
ModuleNotFoundError: No module named 'deal_segregator'
```

Failing test committed alone: `496a8fe test: define deal segregator workbook contract`.

### GREEN

Command: `python3 test_deal_segregator.py && python3 -m py_compile deal_segregator.py`

Result: exit 0:

```text
OK deal segregator engine
```

Engine committed separately: `1d61669 feat: generate segregated deals workbook`.

## Implementation

- Preserved reference parsing, required-column validation, normalization, calculations, sorting, styling, verification metrics.
- Preserved cross-file non-empty Deal-ID deduplication.
- Replaced ZIP/two-workbook output with one XLSX.
- Exact sheet order: `Daily`, `Monthly Summary`, `Verifikasi`.
- Fixture assertions: 9,418 Daily rows; 9,418 Monthly Summary rows; 209,838 unique used rows.
- Duplicate fixture input remains 209,838 used rows.
- CLI default changed to `<input>-hasil.xlsx`.

## Self-review

- Scope: only `test_deal_segregator.py` and `deal_segregator.py` committed; no Flask files changed.
- Reference diff: processing logic unchanged; only output container, docs, import, writer, CLI suffix changed.
- Repository: no tracked `.csv`, `.xlsx`, `.xlsm`, `.zip`, or `.pdf` files.
- Generated test workbooks remain temporary.
- `git diff --check 097a5d6..HEAD`: clean.

## Concerns

- Test depends on the absolute local zern fixture path by plan design.
- Existing ignored financial workbooks and `__pycache__` predate this task; neither tracked nor modified.
- `.superpowers/` remains untracked and contains this requested report plus existing progress metadata.

## Review-finding closure (11 Sep 2026)

### TDD RED

1. Replaced the private absolute fixture with generated repository-portable CSV inputs. Added exact workbook assertions for sheet/header order, Country preservation, Desk derivation, Daily/monthly grouping and sorting, Deals, every financial sum, all Verifikasi metrics, and duplicate workbook values.
2. Added source diagnostics tests for malformed, empty, NaN, and infinite financial values; empty/malformed/impossible Time values (including invalid clock times); ragged rows.
3. Commands/results:
   - `python3 test_deal_segregator.py` → exit 1: `AssertionError: invalid input was accepted` at malformed `Volume`.
   - After initial date validation, `python3 test_deal_segregator.py` → exit 1: `AssertionError: invalid input was accepted` at `Time='2026.07.01 25:00:00'`.
   - `python3 test_deal_segregator.py` → exit 1 on rejected XLSX header: workbook remained open.
   - RED commits: `2200558 test: cover deal input validation and aggregation`; `b0e9f8c test: require xlsx closure on header failure`.

### GREEN

- Financial parser now rejects empty, malformed, NaN, and infinite values with file, row, and column diagnostics before Deals aggregation.
- Time parser now accepts only MT5 timestamps with/without fractional seconds or native Excel dates; rejects empty, malformed, impossible date/time values with file/row/Time diagnostics.
- Ragged rows now fail before indexed access, reporting file, row, expected columns, actual columns.
- Input readers now yield rows instead of materializing raw and filtered full-file lists. This removes the largest redundant copies while preserving proven normalization/aggregation behavior.
- XLSX read-only workbooks close from generator `finally`; output workbook closes after save.
- `write_only=True` deliberately not used: current sheets apply cell styles, number formats, freeze panes, and column widths through random-access worksheet APIs. Converting the writer would be a broader rewrite without demonstrated benefit. The safe reduction targets input copies; aggregate result lists remain necessary for two independently sorted outputs.
- Main GREEN commit: `f795b52 fix: reject invalid deal source values`; workbook-header closure follow-up included in final fix/report commit.

### Verification

- `python3 test_deal_segregator.py && python3 -m py_compile deal_segregator.py test_deal_segregator.py && git diff --check` → exit 0; `OK deal segregator engine`.
- Optional real fixture present. Regression command processed it successfully: `{'file_masuk': 1, 'total': 418913, 'dipakai': 209838, 'harian': 9418, 'bulanan': 9418, 'login_unik': 5071}`.
- No Flask files touched.

### Decisions/concerns

- Blank financial cells are treated as malformed, not zero. Silent zero substitution could corrupt totals.
- Validation applies to used (`Entry = out`) rows. Discarded rows remain irrelevant to report calculations.
- Existing external optional fixture remains untracked and is no longer required by core tests.

## Remaining review findings (11 Sep 2026)

### Empty XLSX — RED

- Test commit: `0f48691 test: require empty xlsx closure`.
- `python3 test_deal_segregator.py` → exit 1 at `next(it)` with uncaught `StopIteration`; the workbook was opened before the failing initialization and no closure path existed.

### Empty XLSX — GREEN

- `_baca_xlsx` now guards all workbook/header initialization exceptions, closes any opened workbook, and exits with a source-specific `BERHENTI: gagal membaca header di '<file>'` diagnostic.
- Focused/full script after fix: `python3 test_deal_segregator.py` → exit 0, `OK deal segregator engine`.

### Streaming aggregation — RED

- Test commit: `e754ee1 test: require streaming deal aggregation`.
- Behavioral lifetime instrumentation wraps five normalized financial values per used row. `python3 test_deal_segregator.py` → exit 1: `AssertionError: normalized rows retained: peak=1000` for 200 input rows.

### Streaming aggregation — GREEN

- Each valid unique `out` row now immediately updates Daily and Monthly dictionaries, Login/Desk sets, Country counters, used count, and Profit total.
- Removed per-file normalized `baris` lists and flattened `semua_baris`; retained only aggregate dictionaries, verification sets/counters, Deal ID set, small per-file summaries, final sorted aggregate lists.
- Dedupe order and semantics, summary labels/values, output ordering, diagnostics, CLI, styled normal workbook writer preserved. No `write_only` rewrite.
- `python3 test_deal_segregator.py && python3 -m py_compile deal_segregator.py test_deal_segregator.py && git diff --check` → exit 0, `OK deal segregator engine`.
- Real fixture parity → `{'file_masuk': 1, 'total': 418913, 'dipakai': 209838, 'harian': 9418, 'bulanan': 9418, 'login_unik': 5071}`.
- An unrelated D&W workbook was initially probed as a fixture and correctly rejected for missing Deal History columns; the original zern fixture was then located from test history and passed exact parity.
- No Flask files changed.
