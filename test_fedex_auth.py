import os

import certifi
import requests


FEDEX_API_KEY = os.environ.get("FEDEX_TEST_API_KEY")
FEDEX_SECRET_KEY = os.environ.get("FEDEX_TEST_SECRET_KEY")

TOKEN_URL = "https://apis-sandbox.fedex.com/oauth/token"


if not FEDEX_API_KEY or not FEDEX_SECRET_KEY:
    raise RuntimeError(
        "Set FEDEX_TEST_API_KEY and FEDEX_TEST_SECRET_KEY "
        "in PowerShell before running this test."
    )


def get_fedex_token() -> str:
    response = requests.post(
        TOKEN_URL,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data={
            "grant_type": "client_credentials",
            "client_id": FEDEX_API_KEY,
            "client_secret": FEDEX_SECRET_KEY,
        },
        timeout=30,
        verify=certifi.where(),
    )

    if not response.ok:
        raise RuntimeError(
            f"FedEx authentication failed with HTTP "
            f"{response.status_code}: {response.text}"
        )

    payload = response.json()
    token = payload.get("access_token")

    if not token:
        raise RuntimeError(
            f"FedEx did not return an access token: {payload}"
        )

    return token


if __name__ == "__main__":
    token = get_fedex_token()
    print("FedEx authentication succeeded.")
    print("Token preview:", token[:20] + "...")