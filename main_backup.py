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
import subprocess
import time
from colorama import Fore, Style, init
import sys
from pathlib import Path
import json
from io import BytesIO



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
    cutoff = now - timedelta(hours=28)
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


def send_email_with_attachment(subject, body, recipient, attachment_path):

    sender_email = os.getenv("SENDER_EMAIL")    
    access_token = get_access_token()
    if not access_token:
        print("Unable to send summary email: No access token.")
        return
    
    with open(attachment_path, "rb") as file:
        file_content = file.read()
    filename = os.path.basename(attachment_path)
    encoded_content = base64.b64encode(file_content).decode("utf-8")

    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": recipient}}],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": filename,
                "contentBytes": encoded_content
            }]
        },
        "saveToSentItems": "true"
    }

    url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, data=json.dumps(message))
    if response.status_code != 202:
        raise Exception(f"Failed to send email: {response.status_code} - {response.text}")





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
            result1 = subprocess.run(["python3", "processor.py", file_path], capture_output=True, text=True)
            print(result1.stdout)

            print("⏳ Waiting 5 seconds for Dropbox to sync...")
            time.sleep(5)

            print("Running DFS_MASTER_LIST_CLOUD.py...")
            result2 = subprocess.run([sys.executable, "DFS_MASTER_LIST_CLOUD.py"], capture_output=True, text=True)

            if result2.returncode == 0:
                print("✅ DFS_MASTER_LIST_CLOUD.py executed successfully.")
                # Locate latest final file
                LISTS_DIR = Path("/Users/yosi/MYTECH Dropbox/Yosef Kroitoro/AUTODFSMAILER/LISTS")
                latest_file = max(LISTS_DIR.glob("DELL_LIST_GRADE_A_*.xlsx"), key=os.path.getctime)
                # Send email with that file
                try:
                    send_email_with_attachment(
                        subject="Your Final GRADE A File",
                        body="Attached is the latest REFURBISHED GRADE A SYSTEMS FILE.",
                        recipient=os.getenv("RECIPIENT_EMAIL"),
                        attachment_path=str(latest_file)
                    )
                    print(f"✅ Final file emailed to {os.getenv('RECIPIENT_EMAIL')}: {latest_file.name}")
                except Exception as e:
                    print(f"❌ Failed to email final file: {e}")
            else:
                print("❌ Error running DFS_MASTER_LIST_CLOUD.py:")
                print(result2.stderr)
    else:
        print("No matching email found in last 24 hours.")
        
        

    send_summary_email(success=success, found=email_found, file_saved=file_saved, subject_line=subject_line)


if __name__ == "__main__":
    main()



##            try:
##                final_file_path = run_master_list_process()
##                print(f"📤 File generated at: {final_file_path}")
##
##                time.sleep(3)
##
##                # Now use final_file_path to attach it in your email:
##                # Dropbox client using refresh token
##                DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN")
##                DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY")
##                DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET")
##
##                dbx = dropbox.Dropbox(
##                    oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
##                    app_key=DROPBOX_APP_KEY,
##                    app_secret=DROPBOX_APP_SECRET
##                )
##
##                
##                
##                print(f"⛳️ Attempting to download from: {final_file_path}")
##                metadata, res = dbx.files_download(final_path)
##
##                # Save
##                with open("./LISTS/temp_downloaded.xlsx", "wb") as f:
##                    f.write(res.content)
##                print("✅ File downloaded locally.")
##                
####                local_path = "./LISTS/temp_downloaded.xlsx"
####                with open(local_path, "wb") as f:
####                    metadata, res = dbx.files_download(final_file_path)
####                    f.write(res.content)
##
##                send_email_with_attachment(
##                    subject="DELL LIST READY",
##                    body="Attached is the final Excel file.",
##                    recipient=RECIPIENT_EMAIL,
##                    attachment_path=local_path
##                )
##
##            except Exception as e:
##                    print(f"❌ Error while generating or sending file: {e}")

##            result2 = subprocess.run([sys.executable, "DFS_MASTER_LIST_CLOUD.py"], capture_output=True, text=True)
##
##            if result2.returncode == 0:
##                print("✅ DFS_MASTER_LIST_CLOUD.py executed successfully.")
##                # Locate latest final file
##                LISTS_DIR = Path("/Users/yosi/MYTECH Dropbox/Yosef Kroitoro/AUTODFSMAILER/LISTS")
##                latest_file = max(LISTS_DIR.glob("DELL_LIST_GRADE_A_*.xlsx"), key=os.path.getctime)
##                # Send email with that file
##                try:
##                    send_email_with_attachment(
##                        subject="Your Final GRADE A File",
##                        body="Attached is the latest REFURBISHED GRADE A SYSTEMS FILE.",
##                        recipient=os.getenv("RECIPIENT_EMAIL"),
##                        attachment_path=str(latest_file)
##                    )
##                    print(f"✅ Final file emailed to {os.getenv('RECIPIENT_EMAIL')}: {latest_file.name}")
##                except Exception as e:
##                    print(f"❌ Failed to email final file: {e}")
##            else:
##                print("❌ Error running DFS_MASTER_LIST_CLOUD.py:")
##                #print(result2.stderr)
