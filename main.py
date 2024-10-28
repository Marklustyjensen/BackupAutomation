#######################################
#                                     #
#  Script created by Mark L. Jensen   #
#  marklustyjensen@gmail.com          #
#  www.marklustyjensen.com            #
#                                     #
#######################################

# Import the necessary libraries
import os
import os.path

# Import the necessary packages
import zipfile
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# Get the current date in YYYY-MM-DD format
current_date = datetime.now().strftime("%Y-%m-%d")

def validation():
    # Define the scopes
    SCOPES = ['https://www.googleapis.com/auth/drive']

    # Setting the credentials to None to check if the credentials are already present
    global creds
    creds = None

    # Checking if the token.json file is present
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # Checking if the credentials are valid
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

# Define the function to archive the files
def archive_file(folder_path):
    # Get the path to the Documents folder
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
    
    # Create the name of the zip file
    zip_filename = os.path.join(documents_folder, f"{current_date}.zip")
    
    # Create a zip file to store the backup
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Walk through the folder and add the files to the zip file
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                # Getting the full paths to the files to be added to the zip file
                file_path = os.path.join(root, file)
                # Checking if the file is not the zip file itself
                if file_path != f"{folder_path}{current_date}":
                    zipf.write(file_path, os.path.relpath(file_path, folder_path))

# Define the function to upload the file to Google Drive
def upload_file():

    validation()

    try:
        # Building the service
        service = build('drive', 'v3', credentials=creds)

        # Checking if the folder is already present
        response = service.files().list(
            q = "name = 'Back-up' and mimeType = 'application/vnd.google-apps.folder'",
            spaces = 'drive'
        ).execute()

        # If the folder is not present, create a new folder
        if not response['files']:
            file_metadata = {
                'name': 'Back-up',
                'mimeType': 'application/vnd.google-apps.folder'
            }

            file = service.files().create(body = file_metadata, fields = 'id').execute()

            folder_id = file.get('id')
        
        else:
            folder_id = response['files'][0].get('id')

        # Specify the file to upload
        file = f'{current_date}.zip'

        # Define the file metadata
        file_metadata = {
            'name': file,
            'parents': [folder_id]
        }

        # Upload the file
        media = MediaFileUpload(f'{os.path.expanduser("~")}/Documents/{file}', mimetype='application/octet-stream')
        upload_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()

    except HttpError as error:
        print(f'An error occurred: {error}')

# Define the function to keep the number of files to a maximum of 5 in the Google Drive folder
def keep_files():
    
    validation()

    try:
        # Building the service
        service = build('drive', 'v3', credentials=creds)

        # Checking if the folder is already present
        response = service.files().list(
            q = "name = 'Back-up' and mimeType = 'application/vnd.google-apps.folder'",
            spaces = 'drive'
        ).execute()

        folder_id = response['files'][0].get('id')

        # Get the list of files in the folder
        results = service.files().list(
            q = f"'{folder_id}' in parents",
            spaces = 'drive',
            fields = 'nextPageToken, files(id, name)',
            pageToken = None
        ).execute()

        # Get the list of files
        files = results.get('files', [])

        # Check if the number of files is greater than 5
        if len(files) > 5:
            # Sort the files by name
            files.sort(key = lambda x: x['name'])

            # Delete the oldest file
            service.files().delete(fileId = files[len(files) - 1]['id']).execute()

    except HttpError as error:
        print(f'An error occurred: {error}')

# Define the function to delete the file after uploading
def delete_file(zip_filename):
    # Delete the zip file
    os.remove(zip_filename)

# Defining the folder to be archived
folder_to_zip = f'{os.path.expanduser("~")}/Documents/BackupFolder'
# Calling the function to archive the files
archive_file(folder_to_zip)

# Calling the function to upload the file to Google Drive
upload_file()

# Calling the function to keep the number of files to a maximum of 5 in the Google Drive folder
keep_files()

# Define the path to the zip file
zip_file_path = os.path.join(os.path.expanduser("~"), "Documents", f"{datetime.now().strftime('%Y-%m-%d')}.zip")
# Calling the function to delete the file after uploading
delete_file(zip_file_path)
