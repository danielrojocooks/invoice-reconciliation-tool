import os
import re
import json
import base64
from datetime import datetime
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
TOKEN_FILE = "gmail_token.json"
OUTPUT_DIR = "gmail_invoices"
VENDOR_A_PAYMENTS_FILE = "vendor_a_payments.json"
VENDOR_B_PAYMENTS_FILE = "vendor_b_payments.json"
VENDOR_C_PAYMENTS_FILE = "vendor_c_payments.json"


def _find_client_secret():
    for f in os.listdir("."):
        if f.startswith("client_secret_") and f.endswith(".json"):
            return f
    raise FileNotFoundError("No client_secret_*.json file found in current directory")


def _get_service():
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(_find_client_secret(), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []

    def handle_data(self, data):
        self._parts.append(data)

    def get_text(self):
        return "\n".join(self._parts)


def _strip_html(html):
    s = _HTMLStripper()
    s.feed(html)
    return s.get_text()


def _get_body_text(payload):
    """Return plain text from an email payload; strips HTML as fallback."""
    html_fallback = None
    for part in _iter_parts(payload):
        mime = part.get("mimeType", "")
        data_b64 = part.get("body", {}).get("data", "")
        if not data_b64:
            continue
        text = base64.urlsafe_b64decode(data_b64).decode("utf-8", errors="replace")
        if mime == "text/plain":
            return text
        if mime == "text/html" and html_fallback is None:
            html_fallback = _strip_html(text)
    return html_fallback or ""


def _parse_vendor_a_receipt(body, msg_date):
    """
    Parse a Vendor A payment receipt email body.

    Returns:
        {
            date:     ISO date string,
            total:    float (payment total),
            invoices: [{invoice_number: str, amount: float}, ...]
        }

    Invoice numbers are the 10-digit numeric portion of Vendor A SO numbers
    (e.g. "0022398186" from "SO101-0022398186").
    """
    # ---- payment total ----
    total = None
    m = re.search(
        r"(?:total|payment\s+amount|amount\s+paid|amount\s+due)\s*[:\$]?\s*([\d,]+\.\d{2})",
        body, re.IGNORECASE,
    )
    if m:
        total = float(m.group(1).replace(",", ""))

    # ---- invoice table rows ----
    # After HTML stripping, each row typically lands on one line:
    #   "0022398186   $295.75"  or  "SO101-0022398186  295.75"
    # Strategy: find every 10-digit number, then grab the first dollar amount
    # that appears on the same line (or immediately after it).
    invoices = []
    seen_nums = set()
    # First pass: SO101-XXXXXXXXXX style (most reliable)
    for row_m in re.finditer(
        r"SO\d{3}-(\d{10})[^\n]*?\$?\s*([\d,]+\.\d{2})", body, re.IGNORECASE
    ):
        num, amt = row_m.group(1), float(row_m.group(2).replace(",", ""))
        if num not in seen_nums:
            invoices.append({"invoice_number": num, "amount": amt})
            seen_nums.add(num)

    # Second pass: bare 10-digit numbers not already captured
    for row_m in re.finditer(r"\b(\d{10})\b[^\n]*?\$?\s*([\d,]+\.\d{2})", body):
        num, amt = row_m.group(1), float(row_m.group(2).replace(",", ""))
        if num not in seen_nums:
            invoices.append({"invoice_number": num, "amount": amt})
            seen_nums.add(num)

    # ---- payment date ----
    date_str = msg_date
    dm = re.search(r"\b(\d{1,2})[/\-](\d{1,2})[/\-](20\d{2})\b", body)
    if dm:
        try:
            date_str = datetime.strptime(
                f"{dm.group(1)}/{dm.group(2)}/{dm.group(3)}", "%m/%d/%Y"
            ).date().isoformat()
        except ValueError:
            pass

    return {"date": date_str, "total": total, "invoices": invoices}


def _iter_parts(payload):
    """Recursively yield all message parts (handles nested multipart)."""
    parts = payload.get("parts", [])
    if not parts:
        yield payload
        return
    for part in parts:
        yield from _iter_parts(part)


def fetch_pdf_attachments():
    """
    Search Gmail for emails with PDF attachments received between
    2025-02-01 and 2025-05-31. Download all PDFs to gmail_invoices/.
    Returns list of local file paths downloaded.
    """
    service = _get_service()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    query = (
        "(subject:invoice OR subject:inv OR subject:bill OR subject:order "
        "OR subject:payment OR subject:statement OR subject:receipt) "
        "has:attachment filename:pdf after:2025/02/01 before:2025/06/01"
    )
    print(f"Searching Gmail: {query}")

    downloaded = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "q": query, "maxResults": 500}
        if page_token:
            kwargs["pageToken"] = page_token

        result = service.users().messages().list(**kwargs).execute()
        messages = result.get("messages", [])

        for msg_ref in messages:
            msg = service.users().messages().get(
                userId="me", id=msg_ref["id"], format="full"
            ).execute()

            for part in _iter_parts(msg.get("payload", {})):
                filename = part.get("filename", "")
                if not filename.lower().endswith(".pdf"):
                    continue

                att_id = part.get("body", {}).get("attachmentId")
                if not att_id:
                    # Inline data (rare for PDFs but handle it)
                    data_b64 = part.get("body", {}).get("data", "")
                    if not data_b64:
                        continue
                    data = base64.urlsafe_b64decode(data_b64)
                else:
                    att = service.users().messages().attachments().get(
                        userId="me", messageId=msg_ref["id"], id=att_id
                    ).execute()
                    data = base64.urlsafe_b64decode(att["data"])

                dest = os.path.join(OUTPUT_DIR, filename)
                if os.path.exists(dest):
                    base, ext = os.path.splitext(filename)
                    dest = os.path.join(OUTPUT_DIR, f"{base}_{msg_ref['id']}{ext}")

                with open(dest, "wb") as f:
                    f.write(data)

                print(f"  Downloaded: {dest}")
                downloaded.append(dest)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    print(f"Total PDFs downloaded: {len(downloaded)}")
    return downloaded


def fetch_vendor_a_payment_receipts():
    """
    Search Gmail for Vendor A payment receipt emails (body-only, no PDF needed).
    Parses each body for payment date, total, and constituent invoice numbers.
    Saves results to vendor_a_payments.json.
    Returns list of parsed receipt dicts.
    """
    service = _get_service()

    query = (
        'subject:"Payment Receipt" from:vendor_afood.com '
        "after:2025/02/01 before:2025/06/01"
    )
    print(f"Searching Gmail for Vendor A receipts: {query}")

    receipts = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "q": query, "maxResults": 500}
        if page_token:
            kwargs["pageToken"] = page_token

        result = service.users().messages().list(**kwargs).execute()
        messages = result.get("messages", [])

        for msg_ref in messages:
            msg = service.users().messages().get(
                userId="me", id=msg_ref["id"], format="full"
            ).execute()

            headers = {
                h["name"]: h["value"]
                for h in msg.get("payload", {}).get("headers", [])
            }

            # Parse message date from headers
            msg_date = ""
            raw_date = headers.get("Date", "")
            if raw_date:
                try:
                    msg_date = parsedate_to_datetime(raw_date).date().isoformat()
                except Exception:
                    pass

            body = _get_body_text(msg.get("payload", {}))
            receipt = _parse_vendor_a_receipt(body, msg_date)
            receipt["message_id"] = msg_ref["id"]
            receipt["subject"] = headers.get("Subject", "")

            inv_summary = [(r["invoice_number"], r["amount"]) for r in receipt["invoices"]]
            print(f"  {msg_date}  total={receipt['total']}  "
                  f"invoices={inv_summary}  subject={receipt['subject']!r}")
            receipts.append(receipt)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    with open(VENDOR_A_PAYMENTS_FILE, "w") as f:
        json.dump(receipts, f, indent=2)

    print(f"Saved {len(receipts)} Vendor A receipts to {VENDOR_A_PAYMENTS_FILE}")
    return receipts


def _parse_vendor_c_receipt(body, msg_date, subject=""):
    """
    Parse an Vendor C payment receipt email body.

    Returns:
        {
            receipt_number: str | None,   e.g. "#1792-4369" from subject
            amount:         float | None, amount paid
            date:           ISO date str, date paid (falls back to message date)
        }
    """
    # Receipt number — prefer [#NNNN-NNNN] from subject, fall back to INV\d+ in body
    receipt_number = None
    subj_m = re.search(r"\[#([\d-]+)\]", subject)
    if subj_m:
        receipt_number = "#" + subj_m.group(1)
    else:
        rn_m = re.search(r"\b(INV\d+)\b", body, re.IGNORECASE)
        if rn_m:
            receipt_number = rn_m.group(1).upper()

    # Amount paid — look for labeled amount first, then any prominent dollar value
    amount = None
    amt_m = re.search(
        r"(?:amount\s+paid|total\s+paid|total\s+charged|amount\s+due|total)\s*[:\$]?\s*\$?\s*([\d,]+\.\d{2})",
        body, re.IGNORECASE,
    )
    if amt_m:
        amount = float(amt_m.group(1).replace(",", ""))
    else:
        # Fall back to the largest dollar amount in the body
        all_amts = [float(v.replace(",", "")) for v in re.findall(r"\$\s*([\d,]+\.\d{2})", body)]
        if all_amts:
            amount = max(all_amts)

    # Date paid — prefer labeled date in body, fall back to message date
    date_str = msg_date
    date_labels = re.search(
        r"(?:date\s+paid|paid\s+on|payment\s+date|date)\s*[:\-]?\s*(\w+\s+\d{1,2},?\s+20\d{2}|\d{1,2}[/\-]\d{1,2}[/\-]20\d{2})",
        body, re.IGNORECASE,
    )
    if date_labels:
        raw = date_labels.group(1).strip()
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%B %d, %Y", "%B %d %Y", "%b %d, %Y"):
            try:
                date_str = datetime.strptime(raw, fmt).date().isoformat()
                break
            except ValueError:
                continue

    return {"receipt_number": receipt_number, "amount": amount, "date": date_str}


def fetch_vendor_c_payment_receipts():
    """
    Search Gmail for Vendor C payment receipt emails (no PDF — body only).
    Subject: "Your Vendor C, Inc. receipt"
    Extracts receipt number, amount paid, and date paid from the email body.
    Saves results to vendor_c_payments.json.
    Returns list of parsed receipt dicts.
    """
    service = _get_service()

    query = (
        'subject:"Your Vendor C, Inc. receipt" '
        "after:2025/02/01 before:2025/06/01"
    )
    print(f"Searching Gmail for Vendor C receipts: {query}")

    receipts = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "q": query, "maxResults": 500}
        if page_token:
            kwargs["pageToken"] = page_token

        result = service.users().messages().list(**kwargs).execute()
        messages = result.get("messages", [])

        for msg_ref in messages:
            msg = service.users().messages().get(
                userId="me", id=msg_ref["id"], format="full"
            ).execute()

            headers = {
                h["name"]: h["value"]
                for h in msg.get("payload", {}).get("headers", [])
            }

            msg_date = ""
            raw_date = headers.get("Date", "")
            if raw_date:
                try:
                    msg_date = parsedate_to_datetime(raw_date).date().isoformat()
                except Exception:
                    pass

            body = _get_body_text(msg.get("payload", {}))
            subject = headers.get("Subject", "")
            receipt = _parse_vendor_c_receipt(body, msg_date, subject)
            receipt["message_id"] = msg_ref["id"]
            receipt["subject"] = subject

            print(f"  {receipt['date']}  receipt={receipt['receipt_number']}  "
                  f"amount={receipt['amount']}  subject={receipt['subject']!r}")
            receipts.append(receipt)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    with open(VENDOR_C_PAYMENTS_FILE, "w") as f:
        json.dump(receipts, f, indent=2)

    print(f"Saved {len(receipts)} Vendor C receipts to {VENDOR_C_PAYMENTS_FILE}")
    return receipts


def _parse_vendor_b_receipt(body, msg_date):
    """
    Parse a Vendor B transaction receipt email body.

    Returns:
        {
            date:     ISO date string (from "Date/Time" field or message date),
            amount:   float | None  (from "Transaction Amount" field),
            order_id: str | None    (from "Order ID" field),
        }
    """
    # Date/Time field — e.g. "Date/Time  03/14/2025 10:22 AM"
    date_str = msg_date
    dm = re.search(
        r"Date[/\s]*Time\s*[:\-]?\s*(\d{1,2}[/\-]\d{1,2}[/\-]20\d{2})",
        body, re.IGNORECASE,
    )
    if dm:
        raw = dm.group(1).replace("-", "/")
        try:
            date_str = datetime.strptime(raw, "%m/%d/%Y").date().isoformat()
        except ValueError:
            pass

    # Transaction Amount — e.g. "Transaction Amount : $$1,152.91" (double-$ in body)
    amount = None
    am = re.search(
        r"Transaction\s+Amount\s*[:\-]?\s*\$+\s*([\d,]+\.\d{2})",
        body, re.IGNORECASE,
    )
    if am:
        amount = float(am.group(1).replace(",", ""))

    # Order ID — e.g. "Order ID  6275436"
    order_id = None
    om = re.search(r"Order\s+ID\s*[:\-]?\s*(\S+)", body, re.IGNORECASE)
    if om:
        order_id = om.group(1).strip()

    return {"date": date_str, "amount": amount, "order_id": order_id}


def fetch_vendor_b_payment_receipts():
    """
    Search Gmail for Vendor B transaction receipt emails (body only).
    Subject: "The Vendor B Transaction Receipt"
    Extracts date, transaction amount, and order ID from the email body.
    Saves results to vendor_b_payments.json.
    Returns list of parsed receipt dicts.
    """
    service = _get_service()

    query = (
        'subject:"The Vendor B Transaction Receipt" '
        "after:2025/02/01 before:2025/06/01"
    )
    print(f"Searching Gmail for Vendor B receipts: {query}")

    receipts = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "q": query, "maxResults": 500}
        if page_token:
            kwargs["pageToken"] = page_token

        result = service.users().messages().list(**kwargs).execute()
        messages = result.get("messages", [])

        for msg_ref in messages:
            msg = service.users().messages().get(
                userId="me", id=msg_ref["id"], format="full"
            ).execute()

            headers = {
                h["name"]: h["value"]
                for h in msg.get("payload", {}).get("headers", [])
            }

            msg_date = ""
            raw_date = headers.get("Date", "")
            if raw_date:
                try:
                    msg_date = parsedate_to_datetime(raw_date).date().isoformat()
                except Exception:
                    pass

            body = _get_body_text(msg.get("payload", {}))
            receipt = _parse_vendor_b_receipt(body, msg_date)
            receipt["message_id"] = msg_ref["id"]
            receipt["subject"] = headers.get("Subject", "")

            print(f"  {receipt['date']}  amount={receipt['amount']}  "
                  f"order_id={receipt['order_id']}  subject={receipt['subject']!r}")
            receipts.append(receipt)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    with open(VENDOR_B_PAYMENTS_FILE, "w") as f:
        json.dump(receipts, f, indent=2)

    print(f"Saved {len(receipts)} Vendor B receipts to {VENDOR_B_PAYMENTS_FILE}")
    return receipts


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--vendor-a-receipts", action="store_true",
                        help="Fetch Vendor A payment receipt emails (body only)")
    parser.add_argument("--vendor-b-receipts", action="store_true",
                        help="Fetch Vendor B transaction receipt emails (body only)")
    parser.add_argument("--vendor-c-receipts", action="store_true",
                        help="Fetch Vendor C payment receipt emails (body only)")
    args = parser.parse_args()

    if args.vendor_a_receipts:
        fetch_vendor_a_payment_receipts()
    elif args.vendor_b_receipts:
        fetch_vendor_b_payment_receipts()
    elif args.vendor_c_receipts:
        fetch_vendor_c_payment_receipts()
    else:
        fetch_pdf_attachments()
