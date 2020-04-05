import os
import io
import tempfile
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import errors
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import shutil


class GoogleDriveHandler:

    SCOPES = ['https://www.googleapis.com/auth/drive']

    def __init__(self):
        self.service = self.__authenticate()
        self.temp_dir = tempfile.mkdtemp()

    def __authenticate(self):
        credentials = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                credentials = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not credentials or not credentials.valid:
            if credentials and credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.SCOPES)
                credentials = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(credentials, token)

        return build('drive', 'v3', credentials=credentials)

    def update_file(self, google_file, updated_file_path):
        print("updating google document...")
        try:
            # First retrieve the file from the API.
            file_id = google_file.get('id')
            file = self.service.files().get(fileId=file_id).execute()

            del file['id']

            # File's new content.
            media_body = MediaFileUpload(updated_file_path)

            # Send the request to the API.
            updated_file = self.service.files().update(
                fileId=file_id,
                body=file,
                media_body=media_body).execute()

            return updated_file

        except errors.HttpError as error:
            print('An error occurred: %s' % error)
            return None

    def download_file(self, google_file):
        request = self.service.files().get_media(fileId=google_file.get('id'))

        file_path = os.path.join(self.temp_dir, google_file.get('name'))
        file_io = io.FileIO(file_path, mode='wb')
        downloader = MediaIoBaseDownload(file_io, request)

        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print("Download progress: %d%%." % int(status.progress() * 100))

        print("Downloaded temporary file to %s" % file_path)

        return file_path

    def download_google_document(self, google_file, file_type, file_ext):
        request = self.service.files().export_media(fileId=google_file.get('id'), mimeType=file_type)

        file_path = os.path.join(self.temp_dir, google_file.get('name') + file_ext)
        file_io = io.FileIO(file_path, mode='wb')
        downloader = MediaIoBaseDownload(file_io, request)

        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print("Download progress: %d%%." % int(status.progress() * 100))

        print("Downloaded temporary file to %s" % file_path)

        return file_path

    def find_file(self, file_name):
        print("finding file: %s" % file_name)
        for file in self.get_drive_file_list():
            if file.get('name') == file_name:
                return file

    def get_drive_file_list(self):
        print("obtaining drive file list...")
        file_list = []
        page_token = None
        while True:
            response = self.service.files().list(spaces='drive', pageToken=page_token).execute()
            page_token = response.get('nextPageToken', None)
            file_list += response.get('files', [])
            if page_token is None:
                break

        return file_list

    def clean_up(self):
        print("cleaning up...")
        shutil.rmtree(self.temp_dir)
