import os
import re
import shutil
import pdfplumber
from datetime import datetime


_TESSERACT_DEFAULT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def _check_tesseract():
    """Return True if tesseract is on PATH or at the default Windows install location."""
    if shutil.which("tesseract") or os.path.isfile(_TESSERACT_DEFAULT_PATH):
        return True
    print(
        "\n[OCR] Tesseract not found. Install it to enable OCR for image-only PDFs:\n"
        "  Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
        "           (download the installer, add to PATH, then restart your terminal)\n"
        "  macOS:   brew install tesseract\n"
        "  Ubuntu:  sudo apt install tesseract-ocr\n"
    )
    return False


_TESSERACT_AVAILABLE = _check_tesseract()


# Labeled total patterns — match "Total Due: $1,234.56" style lines
_TOTAL_LABELS = (
    r"(?<![A-Za-z])(?:grand\s+)?total(?:\s+(?:due|amount|payable|invoice))?",
    r"amount\s+(?:due|payable)",
    r"balance\s+due",
    r"invoice\s+total",
    r"net\s+(?:total|amount)",
)
_TOTAL_RE = re.compile(
    r"(?:" + "|".join(_TOTAL_LABELS) + r")\s*[:\-]?\s*\$?\s*([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_MONEY_RE = re.compile(r"\$?\s*([\d,]+\.\d{2})")

_MONTHS = (
    "January|February|March|April|May|June|July|August|"
    "September|October|November|December"
)

_DATE_FORMATS = [
    # 01/15/2025 or 01-15-2025
    (re.compile(r"\b(\d{1,2})[/\-](\d{1,2})[/\-](20\d{2})\b"), "%m/%d/%Y"),
    # 2025-01-15
    (re.compile(r"\b(20\d{2})[/\-](\d{1,2})[/\-](\d{1,2})\b"), "%Y/%m/%d"),
    # January 15, 2025 or January 15 2025 (4-digit year)
    (re.compile(
        rf"\b({_MONTHS})\s+(\d{{1,2}}),?\s+(20\d{{2}})\b",
        re.IGNORECASE,
    ), "%B %d %Y"),
    # 27 January 25  (DD Month YY — 2-digit year, prepend "20")
    (re.compile(
        rf"\b(\d{{1,2}})\s+({_MONTHS})\s+(\d{{2}})\b",
        re.IGNORECASE,
    ), "%d %B %Y"),
    # 02/01/25  (MM/DD/YY — 2-digit year; negative lookahead avoids 4-digit years)
    (re.compile(r"\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{2})(?!\d)\b"), "%m/%d/%Y"),
]

# Lines that are unlikely to be a vendor name
_SKIP_LINE_RE = re.compile(
    r"^\s*(?:invoice|receipt|bill|statement|from|date|company|customer|reprint|page|to:|from:|address|"
    r"phone|tel|fax|email|www\.|http|#|no\.|po box|\d)",
    re.IGNORECASE,
)

# Prefixes to strip from known-vendor fallback lines
_REMIT_PREFIX_RE = re.compile(
    r"^\s*(?:kindly\s+)?(?:remit|pay)\s+(?:to\s*:?|payment\s+to\s*:?)\s*",
    re.IGNORECASE,
)

# Lines that look like payment instructions or mailing addresses — discard as vendor
_PAYMENT_INSTRUCTION_RE = re.compile(
    r"mail\s+payment|remit|payment\s+to|please\s+(?:mail|send|remit)|"
    r"make\s+(?:checks?|payment)|send\s+(?:checks?|payment)|"
    r"\b(?:suite|ste|ave|blvd|street|st\.|rd\.|drive|dr\.|lane|ln\.)\b|"
    r"\b[A-Z]{2}\s+\d{5}\b",   # state + zip
    re.IGNORECASE,
)

# Looks like a date string (contains a month name)
_DATE_LIKE_RE = re.compile(rf"\b(?:{_MONTHS})\b", re.IGNORECASE)

# ---------------------------------------------------------------------------
# Known vendor list — customize for your business
# Add any vendor names you want the OCR extractor to recognize and normalize.
# ---------------------------------------------------------------------------
_KNOWN_VENDORS = [
    "Vendor A",
    "Vendor B",
    "Vendor C",
    "Coffee Supplier",
    "Bakery Co",
]
_KNOWN_VENDOR_RE = re.compile(
    r"\b(" + "|".join(re.escape(v) for v in _KNOWN_VENDORS) + r")\b",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Filename vendor hint — matches vendor names embedded in PDF filenames.
# Extend this regex to match your own vendors' naming conventions.
# ---------------------------------------------------------------------------
_FILENAME_VENDOR_RE = re.compile(
    r"(?<![A-Za-z])(VendorA|VendorB|VendorC|Bakery|Supplier)(?![A-Za-z])",
    re.IGNORECASE,
)

# YYYY-MM-DD or MM-DD-YYYY (4-digit year required)
_FILENAME_DATE_RE = re.compile(
    r"(?:"
    r"(20\d{2})[.\-_](\d{1,2})[.\-_](\d{1,2})"  # YYYY-MM-DD
    r"|"
    r"(\d{1,2})[.\-_](\d{1,2})[.\-_](20\d{2})"  # MM-DD-YYYY
    r")"
)


def _extract_amount(text):
    # Prefer a labeled total
    m = _TOTAL_RE.search(text)
    if m:
        return float(m.group(1).replace(",", ""))

    # Fall back to the largest dollar amount on a line that contains a total label
    for line in text.splitlines():
        if re.search(r"\btotal\b|\bdue\b|\bpayable\b", line, re.IGNORECASE):
            amounts = [float(v.replace(",", "")) for v in _MONEY_RE.findall(line)]
            if amounts:
                return max(amounts)

    # Last resort: largest dollar amount in the whole document
    all_amounts = [float(v.replace(",", "")) for v in _MONEY_RE.findall(text)]
    return max(all_amounts) if all_amounts else None


def _extract_date(text):
    for pattern, fmt in _DATE_FORMATS:
        m = pattern.search(text)
        if not m:
            continue
        try:
            if fmt == "%d %B %Y":
                date_str = f"{m.group(1)} {m.group(2)} 20{m.group(3)}"
                return datetime.strptime(date_str, fmt).date().isoformat()
            elif fmt == "%m/%d/%Y" and len(m.group(3)) == 2:
                date_str = f"{m.group(1)}/{m.group(2)}/20{m.group(3)}"
                return datetime.strptime(date_str, fmt).date().isoformat()
            elif fmt == "%B %d %Y":
                date_str = f"{m.group(1)} {m.group(2)} {m.group(3)}"
            elif fmt == "%Y/%m/%d":
                date_str = f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
            else:
                date_str = f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
            return datetime.strptime(date_str, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def _extract_vendor(text):
    candidate = None
    for line in text.splitlines():
        line = line.strip()
        if len(line) < 3:
            continue
        if _SKIP_LINE_RE.match(line):
            continue
        if not re.search(r"[A-Za-z]{2,}", line):
            continue
        if _PAYMENT_INSTRUCTION_RE.search(line):
            continue
        candidate = line
        break

    # If the candidate looks like a date string, discard it
    if candidate and _DATE_LIKE_RE.search(candidate):
        candidate = None

    # If candidate already contains a known vendor, trim trailing noise and return early
    if candidate:
        vm = _KNOWN_VENDOR_RE.search(candidate)
        if vm:
            return candidate[:vm.end()].strip()

    # Either no candidate OR candidate doesn't contain a known vendor name.
    # Try a known-vendor scan of the full text as fallback/override.
    m = _KNOWN_VENDOR_RE.search(text)
    if m:
        for line in text.splitlines():
            if _KNOWN_VENDOR_RE.search(line):
                line = line.strip()
                line = _REMIT_PREFIX_RE.sub("", line).strip()
                vm = _KNOWN_VENDOR_RE.search(line)
                if vm:
                    line = line[vm.start():vm.end()].strip()
                candidate = line
                break

    return candidate


def _vendor_from_filename(path):
    stem = os.path.basename(path)
    m = _FILENAME_VENDOR_RE.search(stem)
    return next((g for g in m.groups() if g), None) if m else None


def _date_from_filename(path):
    stem = os.path.basename(path)
    m = _FILENAME_DATE_RE.search(stem)
    if not m:
        return None
    if m.group(1):  # YYYY-MM-DD
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date().isoformat()
        except ValueError:
            return None
    else:  # MM-DD-YYYY
        try:
            return datetime(int(m.group(6)), int(m.group(4)), int(m.group(5))).date().isoformat()
        except ValueError:
            return None


_JUNK_VENDORS = {"order detail"}


def _is_junk_vendor(vendor):
    """Return True if vendor is a garbled OCR artifact or a known-junk placeholder."""
    if vendor is None:
        return False
    if not vendor[:1].isascii():
        return True
    return vendor.strip().lower() in _JUNK_VENDORS


def _detect_document_type(text):
    if re.search(r"\bsales\s+order\b", text, re.IGNORECASE):
        return "sales_order"
    return "invoice"


def _ocr_pdf(path):
    """Extract text from a scanned PDF using pdf2image + pytesseract."""
    from pdf2image import convert_from_path
    import pytesseract

    # Update these paths to match your local Tesseract/Poppler install
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    poppler_path = r'C:\path\to\poppler\bin'   # update for your machine
    images = convert_from_path(path, dpi=300, poppler_path=poppler_path)
    return "\n".join(pytesseract.image_to_string(img) for img in images)


def extract_invoice_data(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

    if len(text.strip()) < 20:
        if _TESSERACT_AVAILABLE:
            print(f"  [OCR] falling back to OCR for {path}")
            try:
                text = _ocr_pdf(path)
            except Exception as e:
                print(f"  [OCR] failed: {e}")
        else:
            print(f"  [OCR] skipping {path} (Tesseract not installed)")

    vendor = _extract_vendor(text)
    if _is_junk_vendor(vendor):
        vendor = _vendor_from_filename(path)

    # Filename date (user-assigned) is more reliable than OCR when present
    date = _date_from_filename(path) or _extract_date(text)

    return {
        "amount": _extract_amount(text),
        "date": date,
        "vendor": vendor,
        "document_type": _detect_document_type(text),
        "file": path,
        "text": text,
    }
