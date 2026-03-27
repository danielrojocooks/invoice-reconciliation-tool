import os
import logging
import argparse
from datetime import datetime

import yaml

import extractor
import matcher
import reporter
import state
from loader import load_transactions_from_csv, load_transactions_from_api
from qb_attach import upload_file, attach_to_transaction

_CONFIG_FILE = "config.yaml"

logger = logging.getLogger(__name__)


def _setup_logging():
    """Configure root logger: INFO to console + timestamped file in logs/."""
    os.makedirs("logs", exist_ok=True)
    stamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join("logs", f"run_{stamp}.log")

    fmt     = "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"

    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        datefmt=datefmt,
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding="utf-8"),
        ],
    )
    return log_file


def _load_config():
    if not os.path.isfile(_CONFIG_FILE):
        raise FileNotFoundError(
            f"{_CONFIG_FILE} not found. Copy config.yaml and fill in your settings."
        )
    with open(_CONFIG_FILE, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _validate_env(dry_run: bool, from_csv: bool):
    """Abort early with a clear message if required env vars are missing."""
    from dotenv import load_dotenv
    load_dotenv()

    missing = []
    # QB API keys are only needed for live API runs or live attachment
    if not from_csv or not dry_run:
        for key in ("QB_REALM_ID", "QB_ACCESS_TOKEN", "QB_REFRESH_TOKEN"):
            if not os.getenv(key):
                missing.append(key)
    if missing:
        raise EnvironmentError(
            "Missing required environment variables (check your .env file): "
            + ", ".join(missing)
        )


def main(dry_run=True, from_csv=False):
    log_file = _setup_logging()

    cfg = _load_config()
    _validate_env(dry_run, from_csv)

    extractor.configure(cfg)
    matcher.configure(cfg)
    reporter.configure(cfg)

    invoice_dir = cfg.get("invoice_dir", "invoices")

    mode_tag = "[DRY RUN]" if dry_run else "[LIVE]"
    src_tag  = "[CSV]"     if from_csv else "[API]"
    logger.info("--- Invoice Attach Bot %s %s ---", mode_tag, src_tag)
    logger.info("Log file: %s", log_file)

    logger.info("Loading QuickBooks transactions...")
    if from_csv:
        transactions = load_transactions_from_csv(
            cfg.get("qb_csv_path", "transactions.csv"),
            cfg.get("include_categories", []),
        )
    else:
        transactions = load_transactions_from_api()
    logger.info("  %d transactions loaded", len(transactions))

    if not os.path.isdir(invoice_dir):
        logger.error("%s/ not found. Add your invoice PDFs to that folder.", invoice_dir)
        return

    pdf_paths = [
        os.path.join(root, f)
        for root, _, files in os.walk(invoice_dir)
        for f in files
        if f.lower().endswith(".pdf")
    ]
    logger.info("Scanning %s/ — %d PDF(s) found", invoice_dir, len(pdf_paths))

    logger.info("Extracting invoice data...")
    invoices = []
    for path in pdf_paths:
        inv = extractor.extract_invoice_data(path)
        invoices.append(inv)
        logger.info("  %-40s  amount=%s  date=%s  vendor=%s",
                    os.path.basename(path), inv["amount"], inv["date"], inv["vendor"])

    logger.info("Matching invoices to transactions...")
    matched, unmatched_txns, orphan_invoices = matcher._match_invoices(transactions, invoices)

    reporter._print_preview(matched, unmatched_txns, orphan_invoices)

    if dry_run:
        logger.info("[DRY RUN] Nothing written to disk. No QB calls made.")
        return

    reporter._write_audit_package(matched, unmatched_txns, orphan_invoices)

    state.load()
    logger.info("Attaching matched invoices to QuickBooks...")
    for m in matched:
        txn = m["transaction"]
        for inv in m["invoices"]:
            fname = os.path.basename(inv["file"])
            if state.is_done(txn["id"], fname):
                logger.info("  SKIP (already attached) %s -> TXN %s", fname, txn["id"])
                continue
            try:
                attachable_id, sync_token = upload_file(inv["file"])
                attach_to_transaction(attachable_id, sync_token, txn["id"])
                state.mark_done(txn["id"], fname)
                logger.info("  OK  %s -> TXN %s", fname, txn["id"])
            except Exception as e:
                logger.error("  ERR %s: %s", fname, e)

    logger.info("Done.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Invoice Attach Bot")
    parser.add_argument(
        "--from-csv",
        action="store_true",
        help="Load QB transactions from transactions.csv instead of the API",
    )
    parser.add_argument(
        "--live",
        action="store_true",
        help="Write audit package and attach invoices to QB (default: dry run)",
    )
    args = parser.parse_args()
    main(dry_run=not args.live, from_csv=args.from_csv)
