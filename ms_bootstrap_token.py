import os
import json
import webbrowser
import requests
from urllib.parse import urlparse, parse_qs
from http.server import HTTPServer, BaseHTTPRequestHandler
from dotenv import load_dotenv
from datetime import datetime, timedelta

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")
TOKEN_FILE = os.getenv("TOKEN_FILE", "ms_token.json")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
AUTH_URL = f"{AUTHORITY}/oauth2/v2.0/authorize"
TOKEN_URL = f"{AUTHORITY}/oauth2/v2.0/token"

SCOPES = ["offline_access", "https://graph.microsoft.com/.default"]

# Web server to receive auth code
class AuthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        query = parse_qs(urlparse(self.path).query)
        code = query.get("code", [None])[0]
        if code:
            self.server.auth_code = code
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"Auth code received. You can close this window.")
        else:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b" No code received.")

def get_auth_code():
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "response_mode": "query",
        "scope": " ".join(SCOPES),
    }
    url = f"{AUTH_URL}?{requests.compat.urlencode(params)}"
    print("🌐 Opening browser for login...")
    webbrowser.open(url)

    server = HTTPServer(("localhost", 8000), AuthHandler)
    server.handle_request()
    return getattr(server, "auth_code", None)

def exchange_code_for_token(code):
    data = {
        "client_id": CLIENT_ID,
        "scope": " ".join(SCOPES),
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET,
    }
    response = requests.post(TOKEN_URL, data=data)
    if response.ok:
        token = response.json()
        token["expires_at"] = (
            datetime.utcnow() + timedelta(seconds=token["expires_in"])
        ).isoformat()
        with open(TOKEN_FILE, "w") as f:
            json.dump(token, f)
        print("✅ Token saved to", TOKEN_FILE)
    else:
        print("❌ Token request failed:")
        print(response.text)

if __name__ == "__main__":
    code = get_auth_code()
    if code:
        print("🔐 Authorization code received.")
        exchange_code_for_token(code)
    else:
        print("❌ Failed to get authorization code.")
