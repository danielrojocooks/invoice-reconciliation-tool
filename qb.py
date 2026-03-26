import base64
import re
import os
import requests
from dotenv import load_dotenv

load_dotenv()

QB_REALM = os.getenv("QB_REALM_ID")

_QB_BASE = "https://quickbooks.api.intuit.com"
_TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
_ENV_FILE = ".env"


def _current_token():
    return os.getenv("QB_ACCESS_TOKEN")


def _refresh_access_token():
    """Exchange the refresh token for a new access token and update .env in place."""
    client_id = os.getenv("QB_CLIENT_ID")
    client_secret = os.getenv("QB_CLIENT_SECRET")
    refresh_token = os.getenv("QB_REFRESH_TOKEN")

    basic = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    r = requests.post(
        _TOKEN_URL,
        headers={
            "Authorization": f"Basic {basic}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data={"grant_type": "refresh_token", "refresh_token": refresh_token},
    )
    r.raise_for_status()
    tokens = r.json()

    new_access = tokens["access_token"]
    new_refresh = tokens.get("refresh_token", refresh_token)

    # Patch .env so the new tokens survive the next process
    with open(_ENV_FILE, "r") as f:
        content = f.read()
    content = re.sub(r"(?m)^QB_ACCESS_TOKEN=.*$", f"QB_ACCESS_TOKEN={new_access}", content)
    content = re.sub(r"(?m)^QB_REFRESH_TOKEN=.*$", f"QB_REFRESH_TOKEN={new_refresh}", content)
    with open(_ENV_FILE, "w") as f:
        f.write(content)

    os.environ["QB_ACCESS_TOKEN"] = new_access
    os.environ["QB_REFRESH_TOKEN"] = new_refresh
    print("QB token refreshed and .env updated.")
    return new_access


def _make_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/text",
    }


def get_transactions():
    url = f"{_QB_BASE}/v3/company/{QB_REALM}/query"

    query = """
    SELECT Id, TxnDate, TotalAmt, PrivateNote, EntityRef
    FROM Purchase
    WHERE TxnDate >= '2025-01-01' AND TxnDate <= '2025-12-31'
    MAXRESULTS 1000
    """

    token = _current_token()
    r = requests.post(url, headers=_make_headers(token), data=query)

    # Retry once after refreshing if we get a 401
    if r.status_code == 401:
        print("QB access token expired, refreshing...")
        token = _refresh_access_token()
        r = requests.post(url, headers=_make_headers(token), data=query)

    r.raise_for_status()
    data = r.json()

    purchases = data.get("QueryResponse", {}).get("Purchase", [])

    transactions = []
    for p in purchases:
        transactions.append({
            "id": p["Id"],
            "date": p["TxnDate"],
            "amount": float(p["TotalAmt"]),
            "memo": p.get("PrivateNote", ""),
            "vendor": p.get("EntityRef", {}).get("name", ""),
        })

    return transactions
