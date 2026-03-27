"""
Tracks which (transaction_id, filename) pairs have been successfully
attached to QuickBooks so that re-runs skip already-processed files.

State is persisted to state/processed.json.
"""
import os
import json
import logging

logger = logging.getLogger(__name__)

_STATE_DIR  = "state"
_STATE_FILE = os.path.join(_STATE_DIR, "processed.json")

# In-memory set of "txn_id::filename" keys loaded at startup.
_processed: set = set()


def load():
    """Load processed set from disk. Call once at startup."""
    global _processed
    if os.path.isfile(_STATE_FILE):
        try:
            with open(_STATE_FILE, encoding="utf-8") as f:
                _processed = set(json.load(f))
            logger.info("State: %d already-processed attachment(s) loaded", len(_processed))
        except (json.JSONDecodeError, OSError) as e:
            logger.warning("Could not read state file %s: %s — starting fresh", _STATE_FILE, e)
            _processed = set()
    else:
        _processed = set()


def is_done(txn_id: str, filename: str) -> bool:
    return f"{txn_id}::{filename}" in _processed


def mark_done(txn_id: str, filename: str):
    """Record a successful attachment and persist immediately."""
    key = f"{txn_id}::{filename}"
    _processed.add(key)
    _save()


def _save():
    os.makedirs(_STATE_DIR, exist_ok=True)
    with open(_STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(_processed), f, indent=2)
