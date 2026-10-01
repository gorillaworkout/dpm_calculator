# KVB Company Access and D&W Design

**Date:** 2026-10-01
**Status:** Approved in chat

## Goal

Add KVB as a sister-company workspace to the existing web app. KVB initially exposes one tool only: Generate D&W. Its upload runs the proven KVB translation and existing calculation scripts as one job.

The existing DPM user experience and processing behavior are a regression boundary. DPM input rules, routes where compatibility matters, menu tools, processing pipeline, generated files, and calculations must remain unchanged.

## Scope

### Included

- Per-user access to DPM, KVB, or both.
- Company switcher showing DPM and KVB; inaccessible companies remain visible but disabled.
- KVB Generate D&W page.
- One-step KVB pipeline:
  1. `siapkan_kvb.py <upload> -o <translated>`
  2. `isi_template.py <translated> --bulan YYYY-MM -o <template>`
  3. `hitung_dw.py <template> --period YYYY-MM -o <result>`
- Admin controls for company access at account creation and later editing.
- Automatic migration of the current SQLite user table.

### Excluded

- KVB Deal Segregator. Its fields are still under review.
- Any modification to DPM calculation logic.
- Any modification to Deal Segregator logic.
- Shared or persisted KVB rate data outside the uploaded workbook.
- Deleting or recreating the user database.

## Compatibility Boundary

DPM must behave exactly as before this feature:

- Existing DPM URLs remain valid.
- Existing DPM tool definitions and navigation labels remain valid.
- `Generate D&W` continues to run only `isi_template.py` then `hitung_dw.py`.
- `Deal Segregator` remains DPM-only and keeps its current upload behavior.
- DPM validation, report-month handling, opening-balance inputs, async jobs, downloads, and filenames remain unchanged.
- Existing DPM tests must pass without weakening their assertions.

KVB is added through separate configuration and route authorization. The implementation must not insert `siapkan_kvb.py` into the DPM pipeline.

## User Access Model

Add two non-null SQLite columns to `users`:

- `can_dpm INTEGER NOT NULL`
- `can_kvb INTEGER NOT NULL`

Migration uses `ALTER TABLE`; it never drops or recreates `users`:

- Existing administrators receive DPM + KVB.
- Existing non-admin users receive DPM only.
- The bootstrap administrator receives DPM + KVB.

New accounts must receive at least one company. Admins can edit company access later. Account administration permission remains controlled by `is_admin`; company access does not grant admin permission.

A user with no allowed companies is invalid. Server validation rejects creating or updating that state.

## Navigation and Authorization

The sidebar includes a company switcher:

- DPM
- KVB

Both companies remain visible. A company without permission is disabled and cannot be selected.

Company determines visible tools:

- DPM: Generate D&W, Deal Segregator.
- KVB: Generate D&W.

Preserve existing DPM routes for backward compatibility. Add company-qualified KVB routes under `/kvb`, including `/kvb/tool/dw`. Every KVB GET, POST, job-status, and download path must verify `can_kvb`. Existing global login protection remains in place.

Direct access without company permission returns HTTP 403. Hiding a menu item alone is insufficient authorization.

Users with both permissions can switch companies. Users with one permission see the other company disabled. The landing page selects an allowed company; it must never route a user into a company they cannot access.

## KVB Generate D&W

The KVB page uses English user-facing copy. It explains:

- Prepare one KVB Plus `.xlsx`/`.xlsm` workbook.
- Required source sheets: Transfer, D, W, Xero, and Handing Fee/Handling Fee.
- Choose the report month.
- The app translates KVB field and sheet names, prepares the clean D&W template, then calculates the final workbook.
- The uploaded file is never modified.

The page reuses the existing D&W month and opening-balance controls only where the downstream scripts support them. KVB does not expose Deal Segregator.

## Pipeline and Job Handling

The KVB job reuses the current disk-backed async job framework and subprocess isolation. Its script list is separate from DPM:

```text
siapkan_kvb.py -> isi_template.py -> hitung_dw.py
```

For each step:

- Input is the previous step's output.
- Output uses a unique path inside the job directory.
- The report period is passed only to scripts that accept it.
- Existing opening-balance arguments continue to target `hitung_dw.py` only.
- A non-zero subprocess exit stops the job.
- The user sees the failing stage and sanitized process output.
- Intermediate files are not offered for download.
- Final output is a new `.xlsx` file with a KVB-specific filename.

The DPM script list remains exactly:

```text
isi_template.py -> hitung_dw.py
```

## Admin UI

Account creation adds company-access checkboxes:

- DPM
- KVB

At least one is required server-side. The existing user table displays access. Each existing account gains an Edit Access action that updates only `can_dpm` and `can_kvb`; it does not reset the password or alter admin status.

Current protections remain:

- Normal users cannot manage accounts.
- Passwords remain hashed.
- Deleted users lose active-session access immediately.
- No database deletion.

## Error Handling

- Missing company permission: HTTP 403 with a clear English message.
- Missing KVB source sheets: show `siapkan_kvb.py`'s actionable sheet error.
- Invalid report month: reject before creating work.
- Invalid extension: reject before processing.
- Pipeline failure: preserve job failure state across Gunicorn workers.
- Missing/expired job: retain current behavior.

No pipeline error may fall back to DPM processing or silently skip KVB translation.

## Testing

Tests must be written before production changes and demonstrate RED then GREEN.

Required guarantees:

1. Existing DB migration preserves users and passwords.
2. Existing admins become DPM+KVB; existing normal users remain DPM-only.
3. New DPM-only, KVB-only, and dual-access users authenticate.
4. Account creation/update rejects no-company access.
5. DPM-only users receive 403 from KVB routes.
6. KVB-only users receive 403 from DPM tools while retaining permitted shared account routes.
7. Disabled company switch state is rendered correctly.
8. Existing DPM routes, forms, and pipelines remain unchanged.
9. KVB D&W invokes the scripts in exact order with the same report month.
10. KVB pipeline failure stops later stages and exposes a useful error.
11. Existing auth, login, chunked upload, DPM calculator, and Deal Segregator suites remain green.
12. A representative KVB workbook completes end-to-end when a fixture is available.

## Implementation Constraint

The working tree already contains uncommitted user work across auth, UI, scripts, and tests. Implementation must not reset, clean, stash, overwrite, or discard those changes. Before editing each file, inspect its current contents. Keep the diff minimal and additive.
