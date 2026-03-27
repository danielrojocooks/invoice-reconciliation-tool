# Invoice Attach Bot

Automated invoice matching and audit package generator for QuickBooks Online.

Scans a folder of vendor invoice PDFs, extracts amount/date/vendor using text extraction and OCR fallback, matches each invoice to a QuickBooks expense transaction, and produces a fully linked audit package (HTML + XLSX) with every transaction documented.

## What it does

1. Loads QB expense transactions from the API or a CSV export
2. Scans an `invoices/` folder of PDF files, extracting amount, date, and vendor name from each
3. Matches each QB transaction to one or more invoice PDFs using amount, date proximity, and fuzzy vendor name scoring
4. Handles consolidated payments — vendors that issue one payment for multiple invoices are pre-matched via a payments JSON file
5. Produces an `audit_package/` folder containing:
   - `summary.html` — sortable audit table with clickable links to every matched invoice
   - `summary.xlsx` — same data in spreadsheet form with hyperlinks
   - `matched/` — copies of all matched invoice PDFs in one flat folder
   - `orphan_invoices.csv` — PDFs that could not be matched to any transaction

## Module structure

| File | Purpose |
|------|---------|
| `main.py` | Thin orchestrator — loads config, wires modules, runs pipeline |
| `loader.py` | Loads QB transactions from CSV export or API |
| `extractor.py` | Extracts amount/date/vendor from invoice PDFs (text + OCR fallback) |
| `matcher.py` | Matches invoices to transactions; handles batches and consolidated payments |
| `reporter.py` | Writes audit package (HTML, XLSX, orphan CSV) and console preview |
| `state.py` | Idempotency — tracks already-attached files in `state/processed.json` |
| `qb_attach.py` | Uploads invoice PDFs to QBO and links them to expense transactions |
| `gmail_fetcher.py` | Downloads vendor invoice PDFs from Gmail via OAuth |
| `qb.py` / `qb_token.py` | QB OAuth2 token management |

## Setup

```bash
pip install -r requirements.txt
```

For OCR support on scanned PDFs (image-only):
- **Windows**: Install [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) and [Poppler](https://github.com/oschwartz10612/poppler-windows/releases)
- **macOS**: `brew install tesseract poppler`
- **Ubuntu**: `sudo apt install tesseract-ocr poppler-utils`

Update the paths at the top of `extractor.py` if Tesseract/Poppler are not on your PATH.

### QuickBooks & Gmail credentials

1. Copy `.env.example` to `.env` and fill in your OAuth2 credentials
2. On first run, the QB and Gmail auth flows will open a browser window to authorize access
3. Tokens are saved locally to `qb_token.json` and `gmail_token.json` (both gitignored)

## Usage

```bash
# Dry run — preview matches, nothing written to disk
python main.py --from-csv

# Live — write audit_package/, attach invoices to QB
python main.py --from-csv --live
python main.py --live
```

Each run writes a timestamped log to `logs/run_YYYYMMDD_HHMMSS.log`.

## Fetching invoices from Gmail

`gmail_fetcher.py` connects to your Gmail account via OAuth and downloads PDF attachments from vendor emails into `gmail_invoices/` by default.

```bash
# Download all PDF attachments from vendor emails
python gmail_fetcher.py

# Generate vendor payments JSON files used by the consolidated payment pre-pass
python gmail_fetcher.py --vendor-a-receipts   # writes vendor_a_payments.json
python gmail_fetcher.py --vendor-b-receipts   # writes vendor_b_payments.json
python gmail_fetcher.py --vendor-c-receipts   # writes vendor_c_payments.json
```

On first run, Gmail OAuth opens a browser window for authorization. The token is cached to `gmail_token.json` and reused on subsequent runs.

By default `gmail_fetcher.py` saves PDFs to `gmail_invoices/`. Update `invoice_dir` in `config.yaml` to match whichever folder you use.

## QuickBooks attachment

`qb_attach.py` is fully implemented and is called automatically during `--live` runs. It uploads each matched invoice PDF to QBO and links it to the corresponding expense transaction using the Attachable API.

**Important:** This requires Intuit **production** OAuth credentials. Sandbox credentials cannot access production company data. To obtain production credentials, complete app verification at [developer.intuit.com](https://developer.intuit.com).

Attachments that have already been uploaded are tracked in `state/processed.json`. Re-running `--live` skips any invoice that was successfully attached in a prior run.

## Configuration

All settings live in `config.yaml`. Copy and edit it to match your environment — no code changes needed.

```yaml
invoice_dir: "invoices"           # folder containing vendor invoice PDFs
audit_dir: "audit_package"        # output folder
qb_csv_path: "transactions.csv"   # QBO CSV export filename

include_categories:               # only include QB Expense rows matching these
  - "Cost of Goods Sold"
  - "Supplies"

known_qb_vendors:                 # vendors that require a passing vendor score
  - "Vendor A"
  - "Vendor B"

known_invoice_vendors:            # vendor names to recognize in invoice text
  - "Vendor A"
  - "Vendor B"

vendor_aliases:                   # normalize alternate names before matching
  - pattern: "vendor\\s*b"
    canonical: "Vendor B"

vendor_score_strong: 80           # score >= this → High confidence
vendor_score_fuzzy: 50            # score >= this → counts toward Medium confidence
amount_exact_cents: 0.01          # tolerance for exact amount match
date_max_days: 60                 # invoice must be within this many days of QB txn
```

### Consolidated payment vendors

For vendors that send one payment covering multiple invoices, create a payments JSON file and point `vendor_a_payments_file` (or `vendor_b_payments_file`) to it in `config.yaml`. The matcher uses the payment total to find the QB transaction and attaches all individual invoice PDFs.

### Manual overrides (`overrides.json`)

Copy `overrides.example.json` to `overrides.json` (gitignored) and add entries for transactions that can't be auto-matched:

```json
{
  "csv_5": {
    "note": "Cash withdrawal — contractor pay",
    "attachments": ["invoices/receipt.pdf"]
  },
  "csv_12": {
    "note": "Multi-invoice payment — overrides auto-match",
    "replace": true,
    "attachments": ["invoices/inv_a.pdf", "invoices/inv_b.pdf"]
  }
}
```

- `attachments` — additional files to attach alongside any auto-matched invoices
- `replace: true` — discard the auto-matched invoice and use only the files listed here
- `note` — free-text note written to the audit summary

## Match confidence

| Level | Criteria |
|-------|----------|
| **High** | Amount matches exactly + vendor score ≥ 80 + date in range |
| **Medium** | Amount matches + date in range, OR amount matches + fuzzy vendor |
| **Low** | Amount matches only |

Transactions with a known QB vendor name require at least a fuzzy vendor match (score ≥ 50) before a match is accepted.

## CSV export format

If using `--from-csv`, export from QuickBooks Online via:
**Reports → Transaction List by Date → Export to CSV**

Rename the exported file to `transactions.csv` (or set `qb_csv_path` in `config.yaml`).

The loader supports both the current and legacy QBO export column layouts.
