import os
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def authenticate_gdrive():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("Error: credentials.json not found. Please provide your Google Drive API credentials.")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('drive', 'v3', credentials=creds)

def get_service():
    return authenticate_gdrive()

def replace_file_on_gdrive(service, file_path, folder_id, custom_name=None):
    if not service:
        print("Skipping upload due to missing credentials.")
        return

    # Increase socket timeout globally
    import socket
    socket.setdefaulttimeout(300)

    name = custom_name if custom_name else os.path.basename(file_path)
    # 1. Search for existing file in folder
    query = f"name = '{name}' and '{folder_id}' in parents and trashed = false"
    file_id = None
    try:
        results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        if items:
            file_id = items[0]['id']
            print(f"  Existing file found (ID: {file_id}). Updating in place...")
    except Exception as e:
        print(f"  Warning during drive search: {e}")

    # 2. Upload new (Update if existing, Create if not)
    size_bytes = os.path.getsize(file_path)
    size_mb = size_bytes / (1024 * 1024)
    print(f"  Uploading {name} ({size_mb:.2f} MB)...")
    
    media = MediaFileUpload(file_path, chunksize=1024*1024, resumable=True)
    
    if file_id:
        # Update existing
        request = service.files().update(fileId=file_id, media_body=media, fields='id')
    else:
        # Create new
        file_metadata = {'name': name, 'parents': [folder_id]}
        request = service.files().create(body=file_metadata, media_body=media, fields='id')
    
    response = None
    import time
    while response is None:
        retries = 0
        while retries < 5:
            try:
                status, response = request.next_chunk()
                if status:
                    print(f"    - Uploaded {int(status.progress() * 100)}%   ", end='\r')
                break # Success, break out of retry loop
            except Exception as e:
                retries += 1
                print(f"\n    [Network Retry {retries}/5] Recovering from error: {e}", end='\r')
                time.sleep(2 ** retries) # Exponential backoff: 2s, 4s, 8s, 16s...
                if retries >= 5:
                    print(f"\n  Upload failed after 5 retries. Network connection too unstable.")
                    return

    print(f"\n  Drive ID: {response.get('id')} uploaded.")
    return response.get('id')

if __name__ == "__main__":
    # Test script - assumes credentials.json and token.pickle are set
    # Folder ID for user: 1RbYL5fmeL0MhCC-E8CdpXMTEzyBiJH8x
    pass
