import os
import json
import requests
from dotenv import load_dotenv

load_dotenv()

REALM = os.getenv("QB_REALM_ID")
TOKEN = os.getenv("QB_ACCESS_TOKEN")


def upload_file(path):

    print("Uploading:", path)

    url = f"https://quickbooks.api.intuit.com/v3/company/{REALM}/upload"

    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/json"
    }

    metadata = {
        "FileName": os.path.basename(path),
        "ContentType": "application/pdf"
    }

    files = {
        "file_metadata_0": (
            None,
            json.dumps(metadata),
            "application/json"
        ),
        "file_content_0": (
            os.path.basename(path),
            open(path, "rb"),
            "application/pdf"
        )
    }

    r = requests.post(url, headers=headers, files=files)

    print("UPLOAD STATUS:", r.status_code)

    data = r.json()

    print("UPLOAD RESPONSE:", data)

    attachable = data["AttachableResponse"][0]["Attachable"]

    attachable_id = attachable["Id"]
    sync_token = attachable["SyncToken"]

    print("ATTACHABLE ID:", attachable_id)
    print("SYNC TOKEN:", sync_token)

    return attachable_id, sync_token


def attach_to_transaction(attachable_id, sync_token, txn_id):

    print("Attaching file to transaction:", txn_id)

    url = f"https://quickbooks.api.intuit.com/v3/company/{REALM}/attachable"

    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    payload = {
        "Id": attachable_id,
        "SyncToken": sync_token,
        "AttachableRef": [
            {
                "EntityRef": {
                    "type": "Purchase",
                    "value": txn_id
                }
            }
        ]
    }

    r = requests.post(url, headers=headers, json=payload)

    print("ATTACH STATUS:", r.status_code)
    print("ATTACH RESPONSE:", r.text)


if __name__ == "__main__":
    attachable_id, sync_token = upload_file("invoices/test.pdf")
    attach_to_transaction(attachable_id, sync_token, "144")