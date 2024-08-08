from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Path to your service account credentials JSON file
SERVICE_ACCOUNT_FILE = '/home/nineleaps/Documents/Python_Gsheet_automation/credentials.json'

# Scopes required to access Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive']

# Authenticate using the service account
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

# ID of the document you want to share
doc_id = '1YJzO3nYOsjixppD31iZy5EvWhTCeGs3QzF93jab6jZY' 

# Share the document with your email
user_permission = {
    'type': 'user',
    'role': 'reader',
    'emailAddress': 'dhanush.venkataraman@nineleaps.com'  # Replace with your email address
}

drive_service.permissions().create(
    fileId=doc_id,
    body=user_permission,
    fields='id'
).execute()

print("Document shared successfully.")