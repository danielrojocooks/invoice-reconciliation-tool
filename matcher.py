import os
import re
import json
from datetime import date
from itertools import combinations

try:
    from thefuzz import fuzz
    _FUZZY = True
except ImportError:
    _FUZZY = False

# ---------------------------------------------------------------------------
# Module-level config — defaults match config.yaml; overridden by configure()
# ---------------------------------------------------------------------------
_STRONG_VENDOR        = 80
_FUZZY_VENDOR         = 50
_AMOUNT_EXACT_CENTS   = 0.01
_AMOUNT_NEAR_PCT      = 0.005
_AMOUNT_TOLERANCE_PCT = 0.035
_DATE_MAX_DAYS        = 60
_BATCH_MAX_SIZE       = 3
_BATCH_DATE_WINDOW    = 7
VENDOR_A_PAYMENTS_FILE = "vendor_a_payments.json"
VENDOR_B_PAYMENTS_FILE = "vendor_b_payments.json"

_KNOWN_QB_VENDOR_RE = re.compile(
    r"\b(Vendor A|Vendor B|Vendor C|Coffee Supplier|Bakery Co)\b",
    re.IGNORECASE,
)
_VENDOR_ALIASES = [
    (re.compile(r"vendor\s*b", re.IGNORECASE), "Vendor B"),
]

_CONF_RANK = {"High": 3, "Medium": 2, "Low": 1}

# Pattern to extract an invoice number from a PDF filename.
_VENDOR_A_INV_NUM_RE = re.compile(r"INV-(\d+)", re.IGNORECASE)


def configure(cfg):
    """Apply values from a config dict (loaded from config.yaml) at startup."""
    global _STRONG_VENDOR, _FUZZY_VENDOR, _AMOUNT_EXACT_CENTS, _AMOUNT_NEAR_PCT
    global _AMOUNT_TOLERANCE_PCT, _DATE_MAX_DAYS, _BATCH_MAX_SIZE, _BATCH_DATE_WINDOW
    global VENDOR_A_PAYMENTS_FILE, VENDOR_B_PAYMENTS_FILE
    global _KNOWN_QB_VENDOR_RE, _VENDOR_ALIASES

    _STRONG_VENDOR        = cfg.get("vendor_score_strong",  _STRONG_VENDOR)
    _FUZZY_VENDOR         = cfg.get("vendor_score_fuzzy",   _FUZZY_VENDOR)
    _AMOUNT_EXACT_CENTS   = cfg.get("amount_exact_cents",   _AMOUNT_EXACT_CENTS)
    _AMOUNT_NEAR_PCT      = cfg.get("amount_near_pct",      _AMOUNT_NEAR_PCT)
    _AMOUNT_TOLERANCE_PCT = cfg.get("amount_tolerance_pct", _AMOUNT_TOLERANCE_PCT)
    _DATE_MAX_DAYS        = cfg.get("date_max_days",        _DATE_MAX_DAYS)
    _BATCH_MAX_SIZE       = cfg.get("batch_max_size",       _BATCH_MAX_SIZE)
    _BATCH_DATE_WINDOW    = cfg.get("batch_date_window_days", _BATCH_DATE_WINDOW)
    VENDOR_A_PAYMENTS_FILE = cfg.get("vendor_a_payments_file", VENDOR_A_PAYMENTS_FILE)
    VENDOR_B_PAYMENTS_FILE = cfg.get("vendor_b_payments_file", VENDOR_B_PAYMENTS_FILE)

    vendors = cfg.get("known_qb_vendors")
    if vendors:
        pattern = "|".join(re.escape(v) for v in vendors)
        _KNOWN_QB_VENDOR_RE = re.compile(rf"\b({pattern})\b", re.IGNORECASE)

    aliases = cfg.get("vendor_aliases")
    if aliases:
        _VENDOR_ALIASES = [
            (re.compile(a["pattern"], re.IGNORECASE), a["canonical"])
            for a in aliases
        ]


# ---------------------------------------------------------------------------
# Vendor helpers
# ---------------------------------------------------------------------------

def _txn_has_known_vendor(txn):
    target = f"{txn.get('vendor') or ''} {txn.get('memo') or ''}"
    return bool(_KNOWN_QB_VENDOR_RE.search(target))


def _apply_vendor_aliases(inv_vendor, filename):
    """Return inv_vendor rewritten to its canonical name if an alias matches."""
    text = f"{inv_vendor or ''} {os.path.basename(filename)}"
    for pattern, canonical in _VENDOR_ALIASES:
        if pattern.search(text):
            return canonical
    return inv_vendor


# ---------------------------------------------------------------------------
# Matching helpers
# ---------------------------------------------------------------------------

def _amount_ok(inv_amount, txn_amount, tolerance=False):
    """Return (matches, exact).
    tolerance=True: also accept txn up to _AMOUNT_TOLERANCE_PCT above invoice.
    """
    if inv_amount is None:
        return False, False
    if abs(inv_amount - txn_amount) <= _AMOUNT_EXACT_CENTS:
        return True, True
    if inv_amount > 0 and abs(txn_amount - inv_amount) / inv_amount <= _AMOUNT_NEAR_PCT:
        return True, False
    if tolerance and inv_amount > 0 and 0 < (txn_amount - inv_amount) / inv_amount <= _AMOUNT_TOLERANCE_PCT:
        return True, False
    return False, False


def _date_ok(inv_date, txn_date, max_days=None):
    """Invoice date must be 0–max_days before the QB transaction date."""
    if max_days is None:
        max_days = _DATE_MAX_DAYS
    if not inv_date or not txn_date:
        return False
    try:
        inv = date.fromisoformat(inv_date)
        txn = date.fromisoformat(txn_date)
        delta = (txn - inv).days
        return 0 <= delta <= max_days
    except (ValueError, TypeError):
        return False


def _vendor_score(inv_vendor, txn_vendor, txn_memo):
    """Return 0–100 fuzzy match score against QB vendor name."""
    if not inv_vendor or not _FUZZY:
        return 0
    targets = [txn_vendor] if txn_vendor else ([txn_memo] if txn_memo else [])
    if not targets:
        return 0
    return max(fuzz.partial_ratio(inv_vendor.lower(), t.lower()) for t in targets)


def _confidence(vscore, date_matches):
    """
    High   = strong vendor match AND date in range
    Medium = date in range OR fuzzy vendor match
    Low    = neither (amount-only match)
    """
    if vscore >= _STRONG_VENDOR and date_matches:
        return "High"
    if date_matches or vscore >= _FUZZY_VENDOR:
        return "Medium"
    return "Low"


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------

def _diagnose_unmatched(inv, transactions):
    """Return a short string explaining why inv could not be matched."""
    if inv["amount"] is None:
        return "no amount extracted"
    exact = [t for t in transactions if _amount_ok(inv["amount"], t["amount"])[0]]
    if exact:
        parts = []
        for t in exact[:3]:
            do = _date_ok(inv["date"], t["date"])
            vs = _vendor_score(inv["vendor"], t["vendor"], t["memo"])
            issues = []
            if not do:
                if not inv["date"]:
                    issues.append("no date")
                else:
                    try:
                        delta = (date.fromisoformat(t["date"]) - date.fromisoformat(inv["date"])).days
                        issues.append(f"date delta={delta}d")
                    except (ValueError, TypeError):
                        issues.append("unparseable date")
            if vs < _FUZZY_VENDOR:
                issues.append(f"vendor score={vs}<{_FUZZY_VENDOR}")
            label = ", ".join(issues) if issues else "already claimed"
            parts.append(f"TXN {t['id']} ({t['vendor'] or 'no vendor'}): {label}")
        return "exact amount hit — " + "; ".join(parts)
    near = sorted(
        [t for t in transactions if abs(t["amount"] - inv["amount"]) / max(t["amount"], 0.01) <= 0.30],
        key=lambda t: abs(t["amount"] - inv["amount"]),
    )
    if near:
        t = near[0]
        diff_pct = (inv["amount"] - t["amount"]) / t["amount"] * 100
        return f"no exact amount match; nearest TXN {t['id']} ${t['amount']:.2f} ({diff_pct:+.1f}%)"
    return "no transaction within 30% of this amount"


def _diagnose_unmatched_txn(txn, invoices):
    """Return a short string explaining why a QB transaction could not be matched."""
    amt = txn["amount"]
    exact = [inv for inv in invoices if inv["amount"] is not None and abs(inv["amount"] - amt) <= _AMOUNT_EXACT_CENTS]
    if exact:
        parts = []
        for inv in exact[:3]:
            eff = _apply_vendor_aliases(inv.get("vendor"), inv.get("file", ""))
            vs = _vendor_score(eff, txn["vendor"], txn["memo"])
            do = _date_ok(inv["date"], txn["date"])
            issues = []
            if vs < _FUZZY_VENDOR:
                issues.append(f"vendor mismatch (inv={eff or '?'}, score={vs})")
            if not do:
                issues.append("date out of range")
            parts.append(", ".join(issues) if issues else "already claimed by another txn")
        return "exact invoice amount exists — " + "; ".join(parts)
    near = sorted(
        [inv for inv in invoices
         if inv["amount"] is not None
         and abs(inv["amount"] - amt) / max(amt, 0.01) <= 0.30],
        key=lambda inv: abs(inv["amount"] - amt),
    )
    if near:
        inv = near[0]
        diff_pct = (inv["amount"] - amt) / amt * 100
        eff = _apply_vendor_aliases(inv.get("vendor"), inv.get("file", ""))
        return f"nearest invoice ${inv['amount']:.2f} ({diff_pct:+.1f}%) — {eff or os.path.basename(inv['file'])}"
    return "no invoice within 30% of this amount"


# ---------------------------------------------------------------------------
# Vendor-specific receipt pre-passes
# ---------------------------------------------------------------------------

def _load_vendor_a_receipts():
    """Load vendor_a_payments.json. Returns (inv_to_receipt dict, receipts list)."""
    if not os.path.exists(VENDOR_A_PAYMENTS_FILE):
        return {}, []
    with open(VENDOR_A_PAYMENTS_FILE, encoding="utf-8") as f:
        receipts = json.load(f)
    inv_to_receipt = {}
    for r in receipts:
        for entry in r.get("invoices", []):
            num = entry.get("invoice_number") or entry
            inv_to_receipt[str(num)] = r
    return inv_to_receipt, receipts


def _is_vendor_a_invoice(inv):
    """True if this PDF is from Vendor A (by vendor name or filename pattern)."""
    vendor = (inv.get("vendor") or "").lower()
    fname  = os.path.basename(inv.get("file", "")).lower()
    return "vendor a" in vendor or fname.startswith("vendor_a_")


def _load_vendor_b_receipts():
    if not os.path.exists(VENDOR_B_PAYMENTS_FILE):
        return []
    with open(VENDOR_B_PAYMENTS_FILE, encoding="utf-8") as f:
        return json.load(f)


def _is_vendor_b_invoice(inv):
    """True if this PDF is from Vendor B."""
    eff = _apply_vendor_aliases(inv.get("vendor"), inv.get("file", ""))
    return (eff or "").lower() == "vendor b"


# ---------------------------------------------------------------------------
# Core matching
# ---------------------------------------------------------------------------

def _match_invoices(transactions, invoices):
    """
    Returns:
        matched        — list of {transaction, invoices, confidence, batch}
        unmatched_txns — list of {transaction, diag}
        orphan_invoices — list of invoice dicts with no QB match
    """
    used = set()
    matched = []

    # ---- Vendor A receipt pre-pass ----
    _inv_to_receipt, receipts = _load_vendor_a_receipts()
    if receipts:
        used_txn_ids = set()
        for r in receipts:
            if r["total"] is None:
                continue
            best_txn = next(
                (txn for txn in transactions
                 if txn["id"] not in used_txn_ids
                 and abs(txn["amount"] - r["total"]) <= _AMOUNT_EXACT_CENTS),
                None,
            )
            if best_txn is None:
                continue

            avail = [(i, inv) for i, inv in enumerate(invoices)
                     if i not in used and _is_vendor_a_invoice(inv)]

            receipt_amts = [e["amount"] for e in r.get("invoices", []) if e.get("amount") is not None]
            pool = list(avail)
            matched_by_amt = []
            for r_amt in receipt_amts:
                hit = next(
                    ((i, inv) for i, inv in pool
                     if inv["amount"] is not None and abs(inv["amount"] - r_amt) <= _AMOUNT_EXACT_CENTS),
                    None,
                )
                if hit:
                    matched_by_amt.append(hit)
                    pool = [(i, inv) for i, inv in pool if i != hit[0]]

            if len(matched_by_amt) == len(receipt_amts) and matched_by_amt:
                group = matched_by_amt
                conf  = "High"
            else:
                conf  = "Medium"
                group = []
                if r.get("date"):
                    try:
                        r_date = date.fromisoformat(r["date"])
                        group  = [
                            (i, inv) for i, inv in avail
                            if inv.get("date")
                            and abs((date.fromisoformat(inv["date"]) - r_date).days) <= 14
                        ]
                    except (ValueError, TypeError):
                        pass

            if not group:
                continue

            matched.append({
                "transaction": best_txn,
                "invoices":    [inv for _, inv in group],
                "confidence":  conf,
                "batch":       len(group) > 1,
                "vscore":      100,
                "date_ok":     _date_ok(r["date"], best_txn["date"]),
                "vendor_a_receipt": True,
            })
            used.update(i for i, _ in group)
            used_txn_ids.add(best_txn["id"])

    # ---- Vendor B receipt pre-pass ----
    vendor_b_receipts = sorted(_load_vendor_b_receipts(), key=lambda r: r.get("date") or "")
    if vendor_b_receipts:
        receipt_dates = []
        for r in vendor_b_receipts:
            try:
                receipt_dates.append(date.fromisoformat(r["date"]) if r.get("date") else None)
            except (ValueError, TypeError):
                receipt_dates.append(None)

        inv_best_receipt = {}
        for i, inv in enumerate(invoices):
            if i in used or not _is_vendor_b_invoice(inv) or not inv.get("date"):
                continue
            try:
                inv_date = date.fromisoformat(inv["date"])
            except (ValueError, TypeError):
                continue
            best_ri, best_delta = None, float("inf")
            for ri, rd in enumerate(receipt_dates):
                if rd is None:
                    continue
                delta = abs((inv_date - rd).days)
                if delta <= 14 and delta < best_delta:
                    best_delta = delta
                    best_ri    = ri
            if best_ri is not None:
                inv_best_receipt[i] = best_ri

        used_txn_ids_b = set()
        for ri, r in enumerate(vendor_b_receipts):
            if not r.get("amount"):
                continue
            r_date_parsed = None
            try:
                r_date_parsed = date.fromisoformat(r["date"]) if r.get("date") else None
            except (ValueError, TypeError):
                pass

            best_txn = None
            for txn in transactions:
                if txn["id"] in used_txn_ids_b:
                    continue
                if abs(txn["amount"] - r["amount"]) <= _AMOUNT_EXACT_CENTS:
                    if r_date_parsed and txn.get("date"):
                        try:
                            if abs((date.fromisoformat(txn["date"]) - r_date_parsed).days) > 5:
                                continue
                        except (ValueError, TypeError):
                            pass
                    best_txn = txn
                    break
            if best_txn is None:
                continue

            group = [(i, inv) for i, inv in enumerate(invoices)
                     if i not in used and inv_best_receipt.get(i) == ri]
            if not group:
                continue

            near_group = [(i, inv) for i, inv in group
                          if inv.get("amount") is not None
                          and _amount_ok(inv["amount"], r["amount"], tolerance=True)[0]]
            if near_group:
                exact = [(i, inv) for i, inv in near_group
                         if abs(inv["amount"] - r["amount"]) <= _AMOUNT_EXACT_CENTS]
                group = (exact or near_group)[:1]
            else:
                if len(group) == 1:
                    inv_amt = group[0][1].get("amount") or 0
                    if inv_amt > 0 and abs(inv_amt - r["amount"]) / r["amount"] > 0.10:
                        continue

            matched.append({
                "transaction": best_txn,
                "invoices":    [inv for _, inv in group],
                "confidence":  "Medium",
                "batch":       len(group) > 1,
                "vscore":      100,
                "date_ok":     _date_ok(r["date"], best_txn["date"]),
                "vendor_b_receipt": True,
            })
            used.update(i for i, _ in group)
            used_txn_ids_b.add(best_txn["id"])

    # ---- General single-invoice pass ----
    already_matched_ids = {m["transaction"]["id"] for m in matched}
    for txn in sorted(transactions, key=lambda t: t["date"] or ""):
        if txn["id"] in already_matched_ids:
            continue

        best_idx = best_conf = None
        best_vs = 0
        best_do = False
        best_is_invoice = False

        for i, inv in enumerate(invoices):
            if i in used:
                continue
            use_tolerance = _is_vendor_b_invoice(inv)
            amt_ok, amt_exact = _amount_ok(inv["amount"], txn["amount"], tolerance=use_tolerance)
            if not amt_ok:
                continue
            eff_vendor = _apply_vendor_aliases(inv["vendor"], inv["file"])
            vs   = _vendor_score(eff_vendor, txn["vendor"], txn["memo"])
            if _txn_has_known_vendor(txn) and vs < _FUZZY_VENDOR:
                continue
            do   = _date_ok(inv["date"], txn["date"])
            conf = _confidence(vs, do)
            if not amt_exact:
                conf = min(conf, "Medium", key=lambda c: _CONF_RANK[c])
            is_invoice = inv.get("document_type", "invoice") != "sales_order"
            better = (
                best_conf is None
                or _CONF_RANK[conf] > _CONF_RANK[best_conf]
                or (_CONF_RANK[conf] == _CONF_RANK[best_conf] and is_invoice and not best_is_invoice)
            )
            if better:
                best_idx, best_conf, best_vs, best_do, best_is_invoice = i, conf, vs, do, is_invoice

        if best_idx is not None:
            matched.append({
                "transaction": txn,
                "invoices":    [invoices[best_idx]],
                "confidence":  best_conf,
                "batch":       False,
                "vscore":      best_vs,
                "date_ok":     best_do,
            })
            used.add(best_idx)
            continue

        # ---- Batch fallback ----
        available = [
            (i, inv) for i, inv in enumerate(invoices)
            if i not in used
            and inv["amount"] is not None
            and 0 < inv["amount"] < txn["amount"]
        ]
        found_batch = False
        for size in range(2, _BATCH_MAX_SIZE + 1):
            for combo in combinations(available, size):
                idxs = [idx for idx, _ in combo]
                invs = [inv for _, inv in combo]
                vendors = {(inv["vendor"] or "").strip().lower() for inv in invs}
                if len(vendors) > 1 or vendors == {""}:
                    continue
                try:
                    dates = [date.fromisoformat(inv["date"]) for inv in invs if inv["date"]]
                except (ValueError, TypeError):
                    continue
                if len(dates) != size or (max(dates) - min(dates)).days > _BATCH_DATE_WINDOW:
                    continue
                if _txn_has_known_vendor(txn):
                    batch_vs = _vendor_score(
                        _apply_vendor_aliases(invs[0]["vendor"], invs[0]["file"]),
                        txn["vendor"], txn["memo"],
                    )
                    if batch_vs < _FUZZY_VENDOR:
                        continue
                if abs(sum(inv["amount"] for inv in invs) - txn["amount"]) <= _AMOUNT_EXACT_CENTS:
                    matched.append({
                        "transaction": txn,
                        "invoices":    invs,
                        "confidence":  "Medium",
                        "batch":       True,
                    })
                    used.update(idxs)
                    found_batch = True
                    break
            if found_batch:
                break

    # ---- Unmatched QB transactions ----
    matched_txn_ids = {m["transaction"]["id"] for m in matched}
    unmatched_txns  = [
        {"transaction": t, "diag": _diagnose_unmatched_txn(t, invoices)}
        for t in transactions
        if t["id"] not in matched_txn_ids
    ]

    # ---- Orphan invoices ----
    orphan_invoices = [inv for i, inv in enumerate(invoices) if i not in used]

    return matched, unmatched_txns, orphan_invoices
