import os
from msal import ConfidentialClientApplication
import requests
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart




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


def send_summary_email(success, found, file_saved, subject_line):
    sender = os.getenv("EMAIL_SENDER_ACCOUNT")  # Add this to your .env
    recipient_email = os.getenv("RECIPIENT_EMAIL")

    print(f"Sending summary email to: {recipient_email}")


    token = get_access_token()
    if not token:
        print("Failed to get token for summary email.")
        return

    subject = "AUTODFSMAILER RUN SUCCESSFUL" if success else "AUTODFSMAILER RUN FAILED"
    body = f"""
    Run Summary:
    Success: {success}
    Email Found: {found}
    File Saved: {file_saved}
    Subject Line: {subject_line}
    """

    email_msg = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
           "toRecipients": [
    {
        "emailAddress": {
            "address": recipient_email
        }
    }
],

        }
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
        headers=headers,
        json=email_msg
    )

    if response.status_code == 202:
        print("Summary email sent.")
    else:
        print(f"Failed to send summary email. Status code: {response.status_code}")
        print(response.text)





def main():
    email_found = False
    file_saved = False
    subject_line = "No match"
    success = True

    print("Accessing Microsoft Graph...")
    token = get_access_token()
    if not token:
        print("Auth failed.")
        success = False
        subject_line = "Auth failed"
        send_summary_email(success=success, found=email_found, file_saved=file_saved, subject_line=subject_line)
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    print("Looking for matching email...")
    messages = fetch_messages(headers)
    target_msg = find_target_email(messages)

    if target_msg:
        email_found = True
        subject_line = target_msg['subject']
        print(f"Match found: {subject_line}")
        file_path = download_attachment(target_msg["id"], headers)

        if file_path:
            file_saved = True
            print("Running processing script...")
            os.system(f"python3 processor.py {file_path}")
    else:
        print("No matching email found in last 24 hours.")

    send_summary_email(success=success, found=email_found, file_saved=file_saved, subject_line=subject_line)


if __name__ == "__main__":
    main()

