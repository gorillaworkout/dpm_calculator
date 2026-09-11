# Deal Segregator Integration Design

## Goal

Expand `d&w-calculator.gorillaworkout.id` into **Dupoin DPM Tools** with two independent menus:

1. Generate D&W
2. Deal Segregator

The existing D&W calculation and deployment topology must remain unchanged.

## Architecture

Use the existing Flask application in `/Users/bayudarmawan/Documents/Dupoin/DPM Malaysia/DW-Calculator` and its deployed copy at `/home/ubuntu/apps/dw-calculator`.

The application remains one PM2 process named `dw-calculator`, served by Gunicorn on `127.0.0.1:3022`. No new service, port, domain, proxy rule, database, or persistent datastore is needed.

Routes:

- `/`: Dupoin DPM Tools landing page with two cards.
- `/tool/dw`: existing Generate D&W workflow.
- `/tool/segregate`: new Deal Segregator workflow.
- `/job/<id>` and `/job/<id>/download`: shared asynchronous job status and download routes.

## Deal Segregator Input

Accept one or more MT5 Deals History files in `.csv`, `.xlsx`, or `.xlsm` format.

Required columns:

- Deal
- Login
- Group
- Country
- Time
- Type
- Entry
- Symbol
- Volume
- Commission
- Fee
- Swap
- Profit
- Currency

The CSV reader must continue supporting MT5 UTF-16 tab-separated exports plus UTF-8 comma-separated exports.

## Deal Segregator Processing

Preserve the calculation behavior from `/Users/bayudarmawan/Documents/Dupoin/zern`:

1. Pool rows from every uploaded file.
2. Keep only rows whose `Entry` value is `out`, case-insensitively.
3. Remove duplicate non-empty Deal IDs across uploaded files.
4. Derive Desk from Group by removing a leading `real\` prefix.
5. Carry Country through unchanged.
6. Convert Time to a date.
7. Aggregate Daily rows by Login, Country, Desk, Date, Type, Symbol, and Currency.
8. Aggregate Monthly Summary rows by Login, Country, Desk, Month, Type, Symbol, and Currency.
9. Sum Volume, Commission, Fee, Swap, and Profit. Count source rows in Deals.
10. Sort by Login, Country, Desk, period, Type, and Symbol.

## Deal Segregator Output

Return one `.xlsx` workbook, not a ZIP archive.

Required sheets:

1. `Daily`
2. `Monthly Summary`
3. `Verifikasi`

`Verifikasi` records uploaded-file counts, source rows, kept/dropped rows, duplicates, unique logins/desks, populated/empty Country counts, aggregate output row counts, and total Profit.

The downloaded filename uses the uploaded filename stem for up to three files, or a file-count label for larger batches, ending in `-hasil.xlsx`.

## UI

Rename the shared shell and landing-page identity to **Dupoin DPM Tools**. The sidebar and landing page expose both tools.

Use the existing dark responsive design. Keep all user-facing content in English.

The shared upload template behaves by menu:

- Generate D&W: one `.xlsx`/`.xlsm` file, report month, optional J Wallet opening balance, template-download link.
- Deal Segregator: multiple `.csv`/`.xlsx`/`.xlsm` files, no report month, no J Wallet field, no D&W template link.

Each tool page explains:

- what files to prepare;
- what processing occurs;
- what result is returned;
- important limitations and verification guidance.

## Job Handling and Errors

Continue the disk-backed asynchronous job model because uploads may exceed Cloudflare's synchronous request duration and Gunicorn has multiple workers.

Each job's state remains in an atomic JSON marker outside its working directory. Uploaded files and generated results are deleted after collection. Abandoned jobs are swept after six hours.

Validation errors render visibly on the tool page. Processing errors produce a clear message plus captured command output rather than a raw HTTP 500.

The D&W result retains its existing Excel MIME type. Deal Segregator uses the same Excel MIME type for its single workbook.

## Testing

Follow test-first development:

1. Change the existing app test to expect two menus and the Deal Segregator route; confirm failure before implementation.
2. Test Deal Segregator form behavior and multi-file input.
3. Test rejection of missing and unsupported files.
4. Test one real MT5 fixture end-to-end, asserting a valid XLSX with exactly the three required sheets.
5. Test duplicate-file upload, asserting Deal IDs are not double-counted.
6. Run the existing D&W end-to-end checks to prove no regression.
7. Run the Deal Segregator source tests after adapting expected output from ZIP/two workbooks to one workbook/three sheets.

## Deployment and Verification

1. Fingerprint all runtime Python and HTML files locally and remotely.
2. Syntax-check changed Python files.
3. Run both real pipelines locally.
4. Create a timestamped VPS tarball backup excluding `.venv` and caches.
5. Upload only changed/new runtime files.
6. Verify local/remote MD5 parity.
7. Restart `dw-calculator` and run `pm2 save`.
8. Confirm PM2 online status, port `3022`, direct localhost response, public landing-page identity, and both menu routes.
9. Run Deal Segregator through the public URL and verify the downloaded workbook and its three sheets.
10. Re-run the D&W public pipeline when a safe fixture size permits; otherwise retain full local E2E plus remote import, health, hash, and route verification, clearly reporting the fixture limitation.

## Non-Goals

- No change to D&W financial calculations.
- No change to Deal Segregator aggregation rules beyond the approved output-container change.
- No Country lookup or enrichment.
- No authentication, database, saved history, settings, or extra reports.
- No nginx, DNS, domain, or Cloudflare change.
