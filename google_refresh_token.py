from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import json

# If modifying these scopes, delete the previously saved token.json
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def refresh_token():
    creds = None
    if os.path.exists('token.json'):
        print("Deleting old token.json...")
        os.remove('token.json')

    print("Launching OAuth flow in browser...")
    flow = InstalledAppFlow.from_client_secrets_file(
        'credentials.json', SCOPES
    )
    creds = flow.run_local_server(port=0)

    with open('token.json', 'w') as token_file:
        token_file.write(creds.to_json())

    print("✅ New token.json generated successfully!")

if __name__ == '__main__':
    refresh_token()
