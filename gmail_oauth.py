"""
Self-contained Gmail OAuth for QBO Invoice Matcher.

Handles OAuth2 flow for Gmail API access. On first run, opens a browser
for consent. Saves token.json for reuse.

Prerequisites:
  1. Create OAuth credentials at https://console.cloud.google.com/apis/credentials
     - Application type: Desktop app
  2. Enable the Gmail API at https://console.cloud.google.com/apis/library/gmail.googleapis.com
  3. Download the client secret JSON into this directory (any client_secret*.json filename works)
"""

from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

_DIR = Path(__file__).parent.resolve()
_TOKEN_PATH = _DIR / 'token.json'


def _find_client_secret() -> Path:
    """Find the Google OAuth client secret JSON in the project directory."""
    # Check for client_secret*.json first (default Google download name)
    matches = list(_DIR.glob('client_secret*.json'))
    if matches:
        return matches[0]

    # Fall back to credentials.json
    fallback = _DIR / 'credentials.json'
    if fallback.exists():
        return fallback

    raise FileNotFoundError(
        f"No client secret JSON found in {_DIR}\n"
        "Download it from Google Cloud Console (APIs & Services > Credentials > "
        "Desktop app > Download JSON) and place the client_secret_*.json file "
        "in the project directory."
    )


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
            client_secret_path = _find_client_secret()
            print(f"Using client secret: {client_secret_path.name}")
            flow = InstalledAppFlow.from_client_secrets_file(
                str(client_secret_path), SCOPES
            )
            creds = flow.run_local_server(port=0)

        _TOKEN_PATH.write_text(creds.to_json())

    return creds
