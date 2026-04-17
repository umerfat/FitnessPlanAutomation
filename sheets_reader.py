"""
Read client data from Google Sheets using OAuth 2.0 (Desktop app).
Tracks processed clients to avoid reprocessing.
"""

import json
import os

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

import config


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/gmail.send",
]
TOKEN_FILE = os.path.join(os.path.dirname(__file__), "token.json")


def _write_secrets_from_env():
    """In CI, write credentials/token files from environment variables."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    token_json = os.environ.get("GOOGLE_TOKEN_JSON")
    if creds_json and not os.path.exists(config.GOOGLE_CREDENTIALS_FILE):
        with open(config.GOOGLE_CREDENTIALS_FILE, "w") as f:
            f.write(creds_json)
    if token_json and not os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "w") as f:
            f.write(token_json)


def _get_credentials():
    """Get OAuth credentials, refreshing or re-authenticating as needed."""
    _write_secrets_from_env()
    creds = None

    # Load saved token if it exists
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # If no valid credentials, authenticate via browser
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if os.environ.get("CI"):
                raise RuntimeError(
                    "Token expired and cannot re-authenticate in CI. "
                    "Run locally first to refresh token.json, then update the GOOGLE_TOKEN_JSON secret."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                config.GOOGLE_CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Save token for next run
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return creds


def _get_sheet():
    """Authenticate and return the worksheet."""
    creds = _get_credentials()
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(config.GOOGLE_SHEET_ID)
    return spreadsheet.worksheet(config.SHEET_NAME)


def _load_processed() -> set:
    """Load set of processed client keys (timestamp + name)."""
    if os.path.exists(config.PROCESSED_FILE):
        with open(config.PROCESSED_FILE, "r") as f:
            return set(json.load(f))
    return set()


def _save_processed(processed: set):
    """Save processed client keys."""
    with open(config.PROCESSED_FILE, "w") as f:
        json.dump(sorted(processed), f, indent=2)


def _client_key(record: dict) -> str:
    """Generate a unique key for a client record."""
    return f"{record.get('Timestamp', '')}|{record.get('Full Name', '')}"


def get_all_clients() -> list[dict]:
    """Get all client records from the sheet."""
    sheet = _get_sheet()
    return sheet.get_all_records()


def get_new_clients() -> list:
    """Get only unprocessed client records."""
    all_clients = get_all_clients()
    processed = _load_processed()
    new_clients = [c for c in all_clients if _client_key(c) not in processed]
    return new_clients


def mark_processed(client: dict):
    """Mark a client as processed."""
    processed = _load_processed()
    processed.add(_client_key(client))
    _save_processed(processed)


def get_client_by_name(name: str):
    """Find the latest entry for a client by name (case-insensitive)."""
    all_clients = get_all_clients()
    matches = [
        c for c in all_clients
        if c.get("Full Name", "").strip().lower() == name.strip().lower()
    ]
    if not matches:
        return None
    # Return the last match (latest entry in sheet)
    return matches[-1]
