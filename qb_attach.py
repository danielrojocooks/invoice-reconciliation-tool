import os
import json
import logging
import requests
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)

REALM = os.getenv("QB_REALM_ID")
TOKEN = os.getenv("QB_ACCESS_TOKEN")


def _check_response(r, context):
    """Raise a descriptive RuntimeError for non-2xx QB API responses."""
    if r.status_code == 401:
        raise RuntimeError(
            f"QB API 401 Unauthorized ({context}). "
            "Access token is expired or invalid — re-run the OAuth flow."
        )
    if r.status_code == 429:
        raise RuntimeError(
            f"QB API 429 Too Many Requests ({context}). "
            "You've hit the rate limit — wait a minute and retry."
        )
    if r.status_code >= 500:
        raise RuntimeError(
            f"QB API {r.status_code} Server Error ({context}): {r.text[:200]}"
        )
    if not r.ok:
        raise RuntimeError(
            f"QB API {r.status_code} ({context}): {r.text[:200]}"
        )


def upload_file(path):
    logger.info("Uploading: %s", path)

    url = f"https://quickbooks.api.intuit.com/v3/company/{REALM}/upload"
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/json",
    }
    metadata = {
        "FileName": os.path.basename(path),
        "ContentType": "application/pdf",
    }

    with open(path, "rb") as pdf_file:
        files = {
            "file_metadata_0": (None, json.dumps(metadata), "application/json"),
            "file_content_0":  (os.path.basename(path), pdf_file, "application/pdf"),
        }
        r = requests.post(url, headers=headers, files=files)

    logger.debug("upload status: %d", r.status_code)
    _check_response(r, f"upload {os.path.basename(path)}")

    data         = r.json()
    attachable   = data["AttachableResponse"][0]["Attachable"]
    attachable_id = attachable["Id"]
    sync_token   = attachable["SyncToken"]

    logger.debug("attachable_id=%s  sync_token=%s", attachable_id, sync_token)
    return attachable_id, sync_token


def attach_to_transaction(attachable_id, sync_token, txn_id):
    logger.info("Attaching %s to transaction %s", attachable_id, txn_id)

    url = f"https://quickbooks.api.intuit.com/v3/company/{REALM}/attachable"
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    payload = {
        "Id": attachable_id,
        "SyncToken": sync_token,
        "AttachableRef": [
            {
                "EntityRef": {
                    "type": "Purchase",
                    "value": txn_id,
                }
            }
        ],
    }

    r = requests.post(url, headers=headers, json=payload)
    logger.debug("attach status: %d", r.status_code)
    _check_response(r, f"attach to txn {txn_id}")


if __name__ == "__main__":
    attachable_id, sync_token = upload_file("invoices/test.pdf")
    attach_to_transaction(attachable_id, sync_token, "144")
