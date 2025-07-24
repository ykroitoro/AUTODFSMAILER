import webbrowser
import http.server
import threading
import requests

CLIENT_ID = "60aa37fb-0e8f-4410-a021-a46da0544e86"
CLIENT_SECRET = "VmU8Q~35op7Cu6RGMeEYac7X1BUp.-r~zyNJ~b86"
TENANT_ID = "9e2b867e-30cf-46d6-bb0b-62a6e51924e8"
REDIRECT_URI = "http://localhost:8000"
SCOPES = "offline_access Mail.Send"

AUTH_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

code_holder = {}

class CodeHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        if "/?code=" in self.path:
            code = self.path.split("code=")[1].split("&")[0]
            code_holder["code"] = code
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write("<h1>Authorization complete. You can close this tab.</h1>".encode("utf-8"))
        else:
            self.send_error(400, "Missing code")

def start_server():
    server = http.server.HTTPServer(('localhost', 8000), CodeHandler)
    server.handle_request()

def get_authorization_code():
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "response_mode": "query",
        "scope": SCOPES,
    }
    url = f"{AUTH_URL}?{'&'.join([f'{k}={v}' for k, v in params.items()])}"
    print("🌐 Opening browser for authorization...")
    webbrowser.open(url)
    thread = threading.Thread(target=start_server)
    thread.start()
    thread.join()
    return code_holder.get("code")

def exchange_code_for_token(code):
    data = {
        "client_id": CLIENT_ID,
        "scope": SCOPES,
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET,
    }
    response = requests.post(TOKEN_URL, data=data)
    return response.json()

if __name__ == "__main__":
    code = get_authorization_code()
    if code:
        print(f"🔑 Authorization code received: {code}")
        token_response = exchange_code_for_token(code)
        print("🔐 Token response:")
        print(token_response)
    else:
        print("❌ Failed to receive authorization code.")
