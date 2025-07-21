import os
from msal import ConfidentialClientApplication
import requests
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64
import dropbox
from dropbox.files import WriteMode
from dropbox.oauth import DropboxOAuth2FlowNoRedirect
from dropbox.exceptions import AuthError



load_dotenv()

print("TENANT_ID:", os.getenv("TENANT_ID"))
# Configuration
##USER_EMAIL = "bulksales@dellrefurbished.com"
##SUBJECT_KEYWORD = "Available Inventory Notification"
##SAVE_PATH = "DFS_LIST.XLSX"
##SAVE_FOLDER = "/tmp"  # Use /tmp for Railway/Render; it's writable
##SENDER_EMAIL = "yosi@myy-tech.com"
##RECIPIENT_EMAIL = "yosi@myy-tech.com"

USER_EMAIL = os.getenv("USER_EMAIL")
SUBJECT_KEYWORD = os.getenv("SUBJECT_KEYWORD")
#SAVE_PATH = os.path.join(os.getenv("SAVE_FOLDER", "/tmp"), os.getenv("SAVE_PATH", "DFS_LIST.XLSX"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
RECIPIENT_EMAIL = os.getenv("RECIPIENT_EMAIL")
SAVE_FOLDER = os.getenv("SAVE_FOLDER", "/tmp")
SAVE_FILENAME = os.getenv("SAVE_PATH", "DFS_LIST.XLSX")
SAVE_PATH = os.path.join(os.getenv("SAVE_FOLDER"), os.getenv("SAVE_FILENAME"))
DROPBOX_ACCESS_TOKEN = os.getenv("DROPBOX_ACCESS_TOKEN")
app_key = os.getenv("DROPBOX_APP_KEY")
app_secret = os.getenv("DROPBOX_APP_SECRET")
refresh_token = os.getenv("DROPBOX_REFRESH_TOKEN")



def get_access_token():
    app = ConfidentialClientApplication(
        client_id=os.getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{os.getenv('TENANT_ID')}",
        client_credential=os.getenv("CLIENT_SECRET")
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token")

def fetch_messages(headers):
    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/mailFolders/inbox/messages?$top=25"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        messages = data.get("value", [])
        print(f">>> Retrieved {len(messages)} messages from inbox.")
        return messages
    else:
        print(f"Failed to fetch messages. Status code: {response.status_code}")
        print(response.text)
        return []


    

def find_target_email(messages):
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(hours=24)
    print(f"Cutoff time: {cutoff.isoformat()}")

    for msg in messages:
        subject = msg.get("subject", "")
        sender = msg.get("from", {}).get("emailAddress", {}).get("address", "")
        received = datetime.fromisoformat(msg["receivedDateTime"].replace("Z", "+00:00"))

        print(f">>> Checking email from: {sender}")
        print(f"    Subject: {subject}")
        print(f"    Received: {received.isoformat()}")

        if (
            sender.lower() == USER_EMAIL.lower()
            and SUBJECT_KEYWORD.lower() in subject.lower()
            and received >= cutoff
        ):
            print(">>> MATCH FOUND")
            return msg

    print(">>> No matching email found.")
    return None



def download_attachment(message_id, headers):
    attachment_url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/messages/{message_id}/attachments"
    response = requests.get(attachment_url, headers=headers)

    if response.status_code != 200:
        print(f"Failed to fetch attachments. Status code: {response.status_code}")
        print(response.text)
        return None

    attachments = response.json().get("value", [])
    for att in attachments:
        print(f"Found attachment: {att['name']}")
        if att["@odata.type"] == "#microsoft.graph.fileAttachment":
            file_data = att["contentBytes"]
            with open(SAVE_PATH, "wb") as f:
                f.write(base64.b64decode(file_data))
            print(f"Attachment saved to: {SAVE_PATH}")
            upload_to_dropbox(SAVE_PATH, "/AUTODFSMAILER/DFS_LIST.XLSX")
            return SAVE_PATH



    print("No file attachment found.")
    return None


def upload_to_dropbox(local_path, dropbox_path):
    app_key = os.getenv("DROPBOX_APP_KEY")
    app_secret = os.getenv("DROPBOX_APP_SECRET")
    refresh_token = os.getenv("DROPBOX_REFRESH_TOKEN")

    # Initialize Dropbox client using refresh token (no more token expiration!)
    dbx = dropbox.Dropbox(
        app_key=app_key,
        app_secret=app_secret,
        oauth2_refresh_token=refresh_token
    )

    try:
        # Delete existing file if present (optional)
        dbx.files_delete_v2(dropbox_path)
    except dropbox.exceptions.ApiError as e:
        if isinstance(e.error, dropbox.files.DeleteError) and e.error.is_path_lookup() and e.error.get_path_lookup().is_not_found():
            pass  # File doesn't exist—nothing to delete
        else:
            raise  # Re-raise other errors

    with open(local_path, "rb") as f:
        dbx.files_upload(f.read(), dropbox_path, mode=WriteMode("overwrite"))

    print(f"Uploaded to Dropbox: {dropbox_path}")

    


def send_summary_email(success=True, found=False, file_saved=False, subject_line=""):
    import requests

    access_token = get_access_token()
    if not access_token:
        print("Unable to send summary email: No access token.")
        return

    status = "SUCCESS" if success else "FAILED"
    body_content = f"""
    AUTODFSMAILER RUN {status}

    - Email Found: {'Yes' if found else 'No'}
    - File Saved: {'Yes' if file_saved else 'No'}
    - Email Subject: {subject_line if subject_line else 'N/A'}
    """

    message = {
        "message": {
            "subject": f"AUTODFSMAILER RUN {status}",
            "body": {
                "contentType": "Text",
                "content": body_content
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": RECIPIENT_EMAIL
                    }
                }
            ]
        }
    }

    endpoint = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.post(endpoint, headers=headers, json=message)

    if response.status_code == 202:
        print("Summary email sent successfully.")
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

