"""One-time helper to save the EasyPost production API key locally."""
from getpass import getpass
from tracking_service import save_api_key

if __name__ == "__main__":
    print("Paste your EasyPost Production API key.")
    key = getpass("API key: ").strip()
    if not key:
        raise SystemExit("No API key entered.")
    print(f"Saved to: {save_api_key(key)}")