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

## Setup

```bash
pip install -r requirements.txt
```

For OCR support on scanned PDFs (image-only):
- **Windows**: Install [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) and [Poppler](https://github.com/oschwartz10612/poppler-windows/releases)
- **macOS**: `brew install tesseract poppler`
- **Ubuntu**: `sudo apt install tesseract-ocr poppler-utils`

Update the paths at the top of `extract.py` if Tesseract/Poppler are not on your PATH.

### QuickBooks & Gmail credentials

1. Copy `.env.example` to `.env` and fill in your OAuth2 credentials
2. On first run, the QB and Gmail auth flows will open a browser window to authorize access
3. Tokens are saved locally to `qb_token.json` and `gmail_token.json` (both gitignored)

## Usage

```bash
# Dry run — preview matches, nothing written to disk
python main.py

# Dry run using a CSV export instead of the QB API
python main.py --from-csv

# Live — write audit_package/, attach invoices to QB
python main.py --from-csv --live
python main.py --live
```

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

By default `gmail_fetcher.py` saves PDFs to `gmail_invoices/`. If you use that folder instead of `invoices/`, update `INVOICE_DIR` at the top of `main.py`:

```python
INVOICE_DIR = "gmail_invoices"
```

## QuickBooks attachment

`qb_attach.py` is fully implemented and is called automatically during `--live` runs. It uploads each matched invoice PDF to QBO and links it to the corresponding expense transaction using the Attachable API.

**Important:** This requires Intuit **production** OAuth credentials. Sandbox credentials cannot access production company data. To obtain production credentials, complete app verification at [developer.intuit.com](https://developer.intuit.com).

## Configuration

### Vendor list (`extract.py`)

The `_KNOWN_VENDORS` list tells the OCR extractor which vendor names to look for in invoice text. Update it with your own vendors:

```python
_KNOWN_VENDORS = [
    "Vendor A",
    "Vendor B",
    "Bakery Co",
]
```

### Vendor aliases (`main.py`)

If a vendor appears under multiple names in QB or on invoices, add an alias to normalize them before matching:

```python
_VENDOR_ALIASES = [
    (re.compile(r"vendor\s*b", re.IGNORECASE), "Vendor B"),
]
```

### Transaction categories (`main.py`)

Only transactions whose QB account category matches `_INCLUDE_CATEGORIES` are included:

```python
_INCLUDE_CATEGORIES = {"Cost of Goods Sold", "Supplies"}
```

### Consolidated payment vendors (`main.py`)

For vendors that send one payment covering multiple invoices, create a payments JSON file (see `vendor_a_payments.json` for the format) and point `VENDOR_A_PAYMENTS_FILE` to it. The matcher will use the payment total to find the QB transaction and attach all individual invoice PDFs.

### Manual overrides (`overrides.json`)

For transactions that can't be auto-matched, add an entry to `overrides.json`:

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

Transactions with a known QB vendor name (`_KNOWN_QB_VENDOR_RE`) require at least a fuzzy vendor match (score ≥ 50) before a match is accepted.

## CSV export format

If using `--from-csv`, export from QuickBooks Online via:
**Reports → Transaction List by Date → Export to CSV**

Rename the exported file to `transactions.csv` before running, or update `QB_CSV_PATH` at the top of `main.py` to match your filename.

The loader supports both the current and legacy QBO export column layouts.
