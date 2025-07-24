import os
import json
import requests
from dotenv import load_dotenv
from datetime import datetime, timedelta

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")
TOKEN_FILE = os.getenv("TOKEN_FILE", "ms_token.json")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
TOKEN_URL = f"{AUTHORITY}/oauth2/v2.0/token"
SCOPES = ["https://graph.microsoft.com/.default"]

def load_token():
    if not os.path.exists(TOKEN_FILE):
        return None
    with open(TOKEN_FILE, "r") as f:
        return json.load(f)

def save_token(token_data):
    token_data["expires_at"] = (
        datetime.utcnow() + timedelta(seconds=token_data["expires_in"])
    ).isoformat()
    with open(TOKEN_FILE, "w") as f:
        json.dump(token_data, f)

def is_token_expired(token_data):
    if "expires_at" not in token_data:
        return True
    return datetime.utcnow() >= datetime.fromisoformat(token_data["expires_at"])

def refresh_access_token(refresh_token):
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "redirect_uri": REDIRECT_URI,
        "scope": "offline_access https://graph.microsoft.com/.default"
    }
    response = requests.post(TOKEN_URL, data=data)
    if response.ok:
        new_token = response.json()
        save_token(new_token)
        return new_token
    else:
        raise Exception(f"🔐 Refresh failed: {response.status_code} - {response.text}")

def get_access_token():
    token_data = load_token()
    if not token_data:
        raise Exception("❌ No token file found. Please authenticate first.")

    if is_token_expired(token_data):
        print("🔄 Access token expired, refreshing...")
        token_data = refresh_access_token(token_data["refresh_token"])
    return token_data["access_token"]
