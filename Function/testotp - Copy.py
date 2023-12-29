from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import base64
import re
import os

def getOtp():
    # Set up the Gmail API
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    creds = None
    token_path = 'token.json'  # Path to store token

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path)

    if not creds or not creds.valid:
        print(f'Token Expiry: {creds.expiry}')
        if creds and creds.expired and creds.refresh_token:
            try:
              creds.refresh(Request())
              print("Token refreshed successfully.")
            except Exception as e:
              print(f"Error refreshing token: {e}")
              flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
              creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    # Create Gmail API service
    service = build('gmail', 'v1', credentials=creds)

    # Get the latest email
    user_id = 'me'
    label_id = 'INBOX'
    messages = service.users().messages().list(userId=user_id, labelIds=[label_id]).execute()
    latest_message_id = messages['messages'][0]['id']

    # Get the message content
    message = service.users().messages().get(userId=user_id, id=latest_message_id).execute()
    print("message======")
    print(message)

    # Get the message content
    if 'parts' in message['payload']:
        # If there are parts, assume the first part contains the body data
        message_data = message['payload']['parts'][0]['body']['data']
    else:
        # If there are no parts, assume the body data is directly in the message
        message_data = message['payload']['body']['data']

    # Decode the message data
    decoded_data = base64.urlsafe_b64decode(message_data).decode('utf-8')
    print("decode data======")
    print(decoded_data)

    # Extract the verification code using regex
    match = re.search(r'Your OTP\s*for\s*Capitaline\s*Login\s*is\s*(\d{4})', decoded_data)


    print(match)


    if match:
        verification_code = match.group(1)
        print(f'Capitaline Login otp : {verification_code}')
        return verification_code
    else:
        print('No Capitaline Login otp found in the email.')
        return None


