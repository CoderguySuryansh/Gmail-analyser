import os
import pickle
import datetime
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Google API credentials and scope
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/spreadsheets']

# Constants for date range
START_DATE = datetime.datetime(2022, 1, 1, 0, 0, 0)
END_DATE = datetime.datetime(2023, 12, 31, 23, 59, 59)

# Regex pattern for email address matching
EMAIL_REGEX = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'

# Function to authorize Gmail API
def authorize_gmail_api():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

# Function to retrieve email data
def get_email_data(service, email_id):
    results = service.users().messages().list(userId=email_id).execute()
    messages = results.get('messages', [])
    email_data = []
    for message in messages:
        msg = service.users().messages().get(userId=email_id, id=message['id']).execute()
        headers = msg['payload']['headers']
        date_str = [header['value'] for header in headers if header['name'] == 'Date'][0]
        sender_str = [header['value'] for header in headers if header['name'] == 'From'][0]
        content_str = msg['snippet']
        date_obj = datetime.datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z')
        if START_DATE <= date_obj <= END_DATE and re.match(EMAIL_REGEX, sender_str):
            email_data.append((date_obj.strftime('%Y-%m-%d %H:%M:%S'), sender_str, content_str))
    return email_data

# Function to write email data to Google Spreadsheet
def write_to_spreadsheet(email_data):
    creds = ServiceAccountCredentials.from_json_keyfile_name('google_credentials.json', SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open('gmaildata').sheet1

    values = [['Date', 'Sender', 'Content']] + email_data
    sheet.insert_rows(values, 2)

# Main function
def main():
    creds = authorize_gmail_api()
    service = build('gmail', 'v1', credentials=creds)

    email_id = 'joshisuryansh2005@gmail.com'  # Replace with your Gmail ID
    email_data = get_email_data(service, email_id)

    write_to_spreadsheet(email_data)

if __name__ == '__main__':
    main()
