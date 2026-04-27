"""One-time helper. Run locally to get an OAuth refresh token for Drive uploads.

Prereq: the client (folder owner) shared the destination folder with YOUR Google
account as Editor. You sign in with your own account here -- not the client's.

Usage:
  1. In Google Cloud Console, create an OAuth 2.0 Client ID (type: Desktop app).
     Download the JSON, save as scripts/oauth_client.json
  2. python scripts/get_refresh_token.py
  3. Sign in in the browser with YOUR Google account (the one the client shared
     the folder with).
  4. Paste the printed [gdrive] block into .streamlit/secrets.toml
     (or into Streamlit Cloud -> Settings -> Secrets for the deployed app).
"""
import os
import sys

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]
HERE = os.path.dirname(os.path.abspath(__file__))
CLIENT_FILE = os.path.join(HERE, "oauth_client.json")
DEFAULT_FOLDER_ID = "1-ivSbgBMGZJwqZ2hp3r7jH4HkzEXjWe7"

if not os.path.isfile(CLIENT_FILE):
    print(f"ERROR: missing {CLIENT_FILE}")
    print("Download the OAuth Client (Desktop app) JSON from Google Cloud Console")
    print("and save it at the path above. See SETUP_GDRIVE.md.")
    sys.exit(1)

flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, SCOPES)
creds = flow.run_local_server(port=0, prompt="consent", access_type="offline")

if not creds.refresh_token:
    print("ERROR: Google did not return a refresh_token.")
    print("Revoke prior access at https://myaccount.google.com/permissions and re-run.")
    sys.exit(2)

print("\n--- Paste this into .streamlit/secrets.toml ---\n")
print("[gdrive]")
print(f'client_id = "{creds.client_id}"')
print(f'client_secret = "{creds.client_secret}"')
print(f'refresh_token = "{creds.refresh_token}"')
print(f'folder_id = "{DEFAULT_FOLDER_ID}"')
