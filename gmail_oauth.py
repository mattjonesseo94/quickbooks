"""
Self-contained Gmail OAuth for QBO Invoice Matcher.

Handles OAuth2 flow for Gmail API access. On first run, opens a browser
for consent. Saves token.json for reuse.

Prerequisites:
  1. Create OAuth credentials at https://console.cloud.google.com/apis/credentials
     - Application type: Desktop app
  2. Enable the Gmail API at https://console.cloud.google.com/apis/library/gmail.googleapis.com
  3. Download the JSON and save as credentials.json in this directory
"""

import os
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

_DIR = Path(__file__).parent.resolve()
_TOKEN_PATH = _DIR / 'token.json'
_CREDENTIALS_PATH = _DIR / 'credentials.json'


def get_gmail_credentials() -> Credentials:
    """
    Return valid Gmail API credentials.

    First run: opens browser for OAuth consent.
    Subsequent runs: loads saved token, refreshing if expired.
    """
    creds = None

    if _TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(_TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not _CREDENTIALS_PATH.exists():
                raise FileNotFoundError(
                    f"Missing {_CREDENTIALS_PATH}\n"
                    "Download OAuth credentials from Google Cloud Console "
                    "and save as credentials.json in the project directory."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(_CREDENTIALS_PATH), SCOPES
            )
            creds = flow.run_local_server(port=0)

        _TOKEN_PATH.write_text(creds.to_json())

    return creds
