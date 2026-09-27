# excel-consolidation

Consolidates a folder of messy regional Excel exports into one clean master
workbook: normalised columns and types, a per-file validation report, and a
change log of every transformation applied. Nothing is silently dropped —
a file that can't be opened or doesn't match the expected columns is
reported as a categorised exception, and every row a file contributes is
accounted for as either accepted, rejected (with a reason), or a duplicate.

**Everything in this repo is synthetic.** The included generator produces
30 fake regional order workbooks with a fixed random seed, so the whole
pipeline can be demoed and tested without any real client data. No code,
data, or spreadsheet layout here is copied from any client or employer
project.

**Job type this proves:** "Excel data consolidation / merge multiple
spreadsheets / clean and standardise Excel data" — the class of Upwork job
where a client has one export per branch/region/month and wants a single
clean spreadsheet plus a record of what was fixed.

## What it does

1. `generate` — builds a deterministic corpus of 30 synthetic regional
   order workbooks with realistic export mess:
   - 4 different header spellings and column orders across files (e.g.
     `Order ID` / `OrderNo` / `Order #` / `Invoice No`);
   - a merged banner/title row above the real header on some files;
   - blank rows and a totals row embedded in the data on most files;
   - order dates rendered as *text* in 5 different formats
     (`2026-03-14`, `03/14/2026`, `14-Mar-2026`, `March 14, 2026`,
     `14.03.2026`);
   - amounts with currency symbols and thousands separators
     (`$1,234.56`, `€999.00`, `1234.56`);
   - duplicate order records, both within a file and copied verbatim into
     a later file;
   - a sheet named differently per file (`Data` / `Sales` / `Sheet1` /
     `Export` / the region name);
   - a region column missing entirely on some files (region is then
     inferred from the filename);
   - 1 corrupt/unreadable file (not a valid `.xlsx` at all);
   - 1 file with the wrong structure entirely (an inventory snapshot with
     no order-data columns).
2. `consolidate` — reads every workbook in a directory and writes one
   `master.xlsx` with three sheets:
   - **Data** — one row per accepted order: `order_id`, `order_date`
     (normalised to ISO `YYYY-MM-DD`), `customer`, `region`, `product`,
     `quantity`, `unit_price`, `amount`, `sales_rep`, plus `source_file`
     and `source_row` for traceability.
   - **Validation** — a per-file summary (rows in, rows out, rejected,
     duplicates, status), a reconciliation total, every rejected row with
     its reason, and every duplicate removed with which file/row it first
     appeared in.
   - **Change log** — every transformation applied, with its count (e.g.
     `dates_normalized: 343`, `currency_symbols_stripped: 262`,
     `duplicate_rows_removed: 27`).

Unreadable or wrong-structure files are reported as categorised exceptions
in the Validation sheet, never silently skipped. Row counts always
reconcile: **rows in = rows out + rejected + duplicates**, per file and
overall (checked in tests, and printed as `reconciled: True/False` by the
CLI and written into the Validation sheet's total row).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## One-command run

```bash
./run_demo.sh
```

Generates the synthetic workbooks into `data/workbooks/` and writes
`master.xlsx` to `data/output/`. Both directories are git-ignored (they're
generated, not source).

## CLI, step by step

```bash
export PYTHONPATH=src   # or: pip install -e . once a pyproject is added

python -m excel_consolidation generate    --out data/workbooks --seed 7
python -m excel_consolidation consolidate --input data/workbooks --out data/output
```

## Tests

```bash
source .venv/bin/activate
export PYTHONPATH=src
pytest tests -v
```

98 tests: header/date/money/quantity normalisation in isolation, the
generator's own ground truth reconciling with itself, the consolidator's
actual output matching that ground truth *exactly* (file by file and in
aggregate — not just "close"), every one of the 11 rejection categories
present and correctly labelled, the corrupt file and the wrong-structure
file each surfacing as a categorised exception (never `ok`, never
skipped), an idempotent re-run producing identical output, region
inferred correctly from the filename when the column is missing, and a
subprocess smoke test of the CLI itself.

## Sample output

Running `./run_demo.sh` against the seed-7 corpus:

```
Generated 30 synthetic regional workbooks in data/workbooks (seed=7).
Processed 30 workbooks from data/workbooks: 500 rows in, 403 in Data, 70 rejected, 27 duplicates removed.
Reconciled: True
```

**Data sheet (first rows):**

| Order Id | Order Date | Customer        | Region    | Product             | Quantity | Unit Price | Amount   | Source File        |
|----------|------------|------------------|-----------|---------------------|----------|------------|----------|---------------------|
| ORD-1014 | 2026-06-06 | Morgan Quill     | Northgate | Widget Type B       | 30       | 180.95     | 5428.50  | 01_northgate.xlsx   |
| ORD-1007 | 2026-02-22 | Dakota Sorensen  | Northgate | Cable Organizer     | 37       | 34.50      | 1276.50  | 01_northgate.xlsx   |
| ORD-1013 | 2026-06-24 | Avery Kingsley   | Northgate | Conference Table    | 20       | 325.33     | 6506.60  | 01_northgate.xlsx   |

**Validation sheet (excerpt):**

| File                    | Sheet | Status              | Rows In | Rows Out | Rejected | Duplicates | Detail                            |
|-------------------------|-------|----------------------|---------|----------|----------|------------|-------------------------------------|
| 01_northgate.xlsx       | Data  | exception            | 17      | 15       | 2        | 0          |                                      |
| 09_millbrook.xlsx       | Export| ok                   | 14      | 14       | 0        | 0          |                                      |
| 15_corrupt_export.xlsx  |       | unreadable           | 0       | 0        | 0        | 0          | BadZipFile: File is not a zip file  |
| 27_inventory_snapshot.xlsx | Inventory | unmappable_structure | 10 | 0 | 10 | 0 | headers did not match the order-data schema |

**Change log (excerpt):**

| Transformation                | Count |
|--------------------------------|-------|
| banner_rows_skipped            | 8     |
| blank_rows_removed             | 26    |
| currency_symbols_stripped      | 262   |
| dates_normalized               | 343   |
| duplicate_rows_removed         | 27    |
| files_region_inferred_from_filename | 5 |
| files_unmappable_structure     | 1     |
| files_unreadable                | 1     |
| totals_rows_removed            | 14    |

## How consolidation works

- **Header detection** scans the first 6 rows of every sheet in a
  workbook for the row with the most column labels it recognises against
  a fixed set of header spellings (`schema.py`), so a merged title/banner
  row above the real header, or a differently-named sheet, doesn't matter.
  A file where no sheet reaches at least 6 recognised columns is reported
  as `unmappable_structure` — the whole file, every row, none silently
  dropped.
- **Row classification**, in order: an all-blank row is `blank_row`; a row
  carrying a text marker like "Total"/"Grand Total"/"Subtotal" anywhere is
  `totals_row`; a missing required field (`order_id`, `order_date`,
  `product`, `quantity`, `amount`) is `missing_<field>`; an unparsable date
  is `invalid_date`; an unparsable quantity or amount is `invalid_quantity`
  / `invalid_amount`. Only after all of that does a row become a candidate
  for the Data sheet.
- **Duplicates** are detected by `order_id` across the *whole* run (not
  just within one file): the first file/row to introduce an order id wins
  a place in Data; every later occurrence, in the same file or a different
  one, is logged in the Validation sheet's duplicates table with which
  file/row it duplicates.
- **Region fallback**: if a file has no recognisable region column, or a
  row's region cell is blank, the region is inferred from the filename
  (`05_cedar_hollow.xlsx` → `Cedar Hollow`) and the substitution is counted
  in the change log.

## Limits

- Header matching covers the label variants demonstrated here plus a
  short list of substring fallbacks (`schema.py: FIELD_KEYWORDS`) — a
  real client's export with a genuinely novel column vocabulary would need
  its own synonyms added, the same way a new invoice vendor needs its own
  label patterns in the PDF-invoice demo.
- Date parsing covers 5 unambiguous text formats (`schema.py:
  DATE_FORMATS`); a format outside that list is correctly rejected as
  `invalid_date`, not guessed at.
- No currency conversion: amounts are normalised (symbols and thousands
  separators stripped) but not converted between currencies, and this
  corpus doesn't mix currencies within a column.
- Duplicate detection keys on `order_id` alone; two genuinely different
  orders that happen to share an id (a real collision, not a re-export)
  would be treated as a duplicate, same trade-off as the PDF-invoice
  demo's duplicate-invoice-number check.
- `master.xlsx` is rebuilt from scratch on every run; there is no
  incremental "append this month's files to last month's master" mode.

## Project layout

```
src/excel_consolidation/
  regions.py       fictional region/customer/product/rep name pools
  schema.py        canonical columns, header spellings, parsing helpers
                    (shared by the generator and the consolidator)
  models.py        FileManifest / FileResult / RejectedRow / DuplicateRow
                    dataclasses shared by generator + consolidator
  generator.py      builds the deterministic messy corpus and writes it
  consolidator.py   reads the corpus, normalises it, writes master.xlsx
  cli.py            `generate` / `consolidate` subcommands
tests/              pytest suite (normalisation, reconciliation against
                    generator ground truth, failure classes, idempotency,
                    CLI smoke test)
LICENSES.md         every open-source library used and its licence
```

## Role

Built with AI assistance. Role: automation/data engineer — I directed the
build (schema, mess taxonomy, reconciliation invariant, test plan) and
reviewed every file; Claude Sonnet wrote the code and tests under that
direction. All data is synthetic; no client or employer code, data, or
spreadsheet layout was used.
