import os
from msal import ConfidentialClientApplication
import requests
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv

load_dotenv()
print("TENANT_ID:", os.environ.get("TENANT_ID"))
# Configuration
EMAIL_TO_WATCH = "bulksales@dellrefurbished.com"
SUBJECT_KEYWORD = "Available Inventory Notification"
FILENAME_TO_SAVE = "DFS_LIST.XLSX"
SAVE_FOLDER = "/tmp"  # Use /tmp for Railway/Render; it's writable

def get_access_token():
    app = ConfidentialClientApplication(
        client_id=os.getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{os.getenv('TENANT_ID')}",
        client_credential=os.getenv("CLIENT_SECRET")
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token")

def fetch_messages(headers):
    url = f"https://graph.microsoft.com/v1.0/users/{os.getenv('TARGET_USER')}/messages?$top=10&$orderby=receivedDateTime desc"
    response = requests.get(url, headers=headers)
    return response.json().get("value", [])

def find_target_email(messages):
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(hours=24)
    for msg in messages:
        subject = msg["subject"]
        sender = msg["from"]["emailAddress"]["address"]
        received = datetime.fromisoformat(msg["receivedDateTime"].replace("Z", "+00:00"))
        if (EMAIL_TO_WATCH.lower() == sender.lower()
            and SUBJECT_KEYWORD.lower() in subject.lower()
            and received >= cutoff):
            return msg
    return None

def download_attachment(message_id, headers):
    url = f"https://graph.microsoft.com/v1.0/users/{os.getenv('TARGET_USER')}/messages/{message_id}/attachments"
    response = requests.get(url, headers=headers)
    attachments = response.json().get("value", [])

    for att in attachments:
        if att.get("@odata.mediaContentType") and "xlsx" in att.get("name", "").lower():
            content_bytes = att["contentBytes"]
            file_path = os.path.join(SAVE_FOLDER, FILENAME_TO_SAVE)
            with open(file_path, "wb") as f:
                f.write(bytes.fromhex(content_bytes))
            print(f"Saved attachment to {file_path}")
            return file_path
    print("No XLSX attachment found.")
    return None

def main():
    print("Accessing Microsoft Graph...")
    token = get_access_token()
    if not token:
        print("Auth failed.")
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    print("Looking for matching email...")
    messages = fetch_messages(headers)
    target_msg = find_target_email(messages)

    if target_msg:
        print(f"Match found: {target_msg['subject']}")
        file_path = download_attachment(target_msg["id"], headers)

        if file_path:
            print("Running processing script...")
            os.system(f"python3 processor.py {file_path}")
    else:
        print("No matching email found in last 24 hours.")

if __name__ == "__main__":
    main()
