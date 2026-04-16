import os
import json
import gspread
from google.oauth2.service_account import Credentials

READ_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

WRITE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_client_cache = None
_write_client_cache = None


def _build_client(scopes):
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        return None, "GOOGLE_SERVICE_ACCOUNT_JSON secret not set. Add your service account JSON to Replit Secrets."
    try:
        info = json.loads(raw)
    except json.JSONDecodeError as e:
        return None, f"Invalid JSON in GOOGLE_SERVICE_ACCOUNT_JSON: {e}"
    try:
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        client = gspread.authorize(creds)
        return client, None
    except Exception as e:
        return None, f"Google auth failed: {e}"


def get_client():
    global _client_cache
    if _client_cache is not None:
        return _client_cache, None
    client, err = _build_client(READ_SCOPES)
    if not err:
        _client_cache = client
    return client, err


def get_write_client():
    global _write_client_cache
    if _write_client_cache is not None:
        return _write_client_cache, None
    client, err = _build_client(WRITE_SCOPES)
    if not err:
        _write_client_cache = client
    return client, err


def reset_client():
    global _client_cache, _write_client_cache
    _client_cache = None
    _write_client_cache = None


def is_configured():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    return bool(raw)
