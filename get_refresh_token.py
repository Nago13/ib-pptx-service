"""
One-time script to obtain a Google OAuth2 refresh token.

Prerequisites:
  1. Go to Google Cloud Console > APIs & Services > Credentials
  2. Create an OAuth 2.0 Client ID (type: Desktop app)
  3. Download the JSON and save it as 'oauth_client.json' in this folder
     OR pass the path via --client-secrets argument

Usage:
  python get_refresh_token.py
  python get_refresh_token.py --client-secrets path/to/client.json

The script opens your browser for Google login. After authorizing,
it prints the refresh token, client_id, and client_secret that you
need to set as environment variables on Render.
"""

import argparse
import json
import os

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]


def main():
    parser = argparse.ArgumentParser(description="Get Google OAuth2 refresh token")
    parser.add_argument(
        "--client-secrets",
        default="oauth_client.json",
        help="Path to OAuth client secrets JSON (default: oauth_client.json)",
    )
    args = parser.parse_args()

    if not os.path.exists(args.client_secrets):
        print(f"ERROR: File not found: {args.client_secrets}")
        print()
        print("To create it:")
        print("  1. Go to https://console.cloud.google.com/apis/credentials")
        print("  2. Click '+ CREATE CREDENTIALS' > 'OAuth client ID'")
        print("  3. Application type: 'Desktop app'")
        print("  4. Download the JSON and save it here as 'oauth_client.json'")
        return

    flow = InstalledAppFlow.from_client_secrets_file(args.client_secrets, SCOPES)
    creds = flow.run_local_server(port=0)

    with open(args.client_secrets) as f:
        client_data = json.load(f)

    installed = client_data.get("installed", client_data.get("web", {}))
    client_id = installed["client_id"]
    client_secret = installed["client_secret"]

    print()
    print("=" * 60)
    print("  OAuth2 credentials obtained successfully!")
    print("=" * 60)
    print()
    print("Set these environment variables on Render:")
    print()
    print(f"  GOOGLE_CLIENT_ID={client_id}")
    print(f"  GOOGLE_CLIENT_SECRET={client_secret}")
    print(f"  GOOGLE_REFRESH_TOKEN={creds.refresh_token}")
    print()
    print("You can also keep DRIVE_FOLDER_ID if you want files")
    print("created inside a specific folder.")
    print()


if __name__ == "__main__":
    main()
