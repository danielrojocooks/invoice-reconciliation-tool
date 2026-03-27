import csv
import os
from datetime import datetime

from qb import get_transactions  # noqa: F401 — re-exported for main.py


# Column name mappings for both QBO export formats
_COL_MAPS = [
    ("Type",             "Category",         "Total",  "Payee"),   # new format
    ("Transaction type", "Account full name", "Amount", "Name"),    # old format
]


def _parse_date(raw):
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(raw.strip(), fmt).date().isoformat()
        except ValueError:
            continue
    return ""


def _parse_amount(raw):
    """Strip $, commas, parens; return float or None."""
    raw      = raw.strip().replace(",", "").replace("$", "")
    negative = raw.startswith("(") and raw.endswith(")")
    raw      = raw.strip("()")
    try:
        return -float(raw) if negative else float(raw)
    except ValueError:
        return None


def load_transactions_from_csv(path, include_categories):
    """Load QB expense transactions from a CSV export.

    Args:
        path: path to the QBO CSV export file
        include_categories: set/list of account category strings to include

    Returns:
        list of transaction dicts with keys: id, date, amount, vendor, memo
    """
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader     = csv.DictReader(f)
        rows       = list(reader)
        fieldnames = reader.fieldnames or []

    col_type = col_cat = col_amount = col_vendor = None
    for type_col, cat_col, amt_col, vendor_col in _COL_MAPS:
        if type_col in fieldnames:
            col_type, col_cat, col_amount, col_vendor = type_col, cat_col, amt_col, vendor_col
            break

    if col_type is None:
        raise ValueError(f"Unrecognised CSV format — headers: {fieldnames}")

    transactions = []
    seq = 1
    for row in rows:
        if row.get(col_type, "").strip() != "Expense":
            continue
        category = row.get(col_cat, "").strip()
        if not any(cat in category for cat in include_categories):
            continue
        raw_amount = _parse_amount(row.get(col_amount, ""))
        if raw_amount is None:
            continue
        transactions.append({
            "id":     f"csv_{seq}",
            "date":   _parse_date(row.get("Date", "").strip()),
            "amount": abs(raw_amount),
            "vendor": row.get(col_vendor, "").strip(),
            "memo":   category,
        })
        seq += 1

    return transactions


def load_transactions_from_api():
    """Load QB transactions via the QuickBooks Online API."""
    return get_transactions()
