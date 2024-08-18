import os
import pickle
from flask import Flask, render_template, request
import google.auth.transport.requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

app = Flask(__name__)

# Replace with your actual Google Spreadsheet ID
SPREADSHEET_ID = '1eAZt-gjE_BpciYDZENUwZ5MUI9L7vSjZ16a7yL5vuo8'  # Ensure this is correct
RANGE_NAME = 'Sheet1!A:D'  # Replace 'Sheet1' with the actual sheet name in your spreadsheet

# Path to your client_secret.json file
CLIENT_SECRET_FILE = r'C:\Users\armor\Downloads\CSC6304WebPage\inventory\client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']  # Read-only access to Google Sheets API

def get_creds():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def search_item_in_sheet(spreadsheet_id, item_name):
    creds = get_creds()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=RANGE_NAME).execute()
    values = result.get('values', [])

    # Debug: Print the fetched values
    print("Fetched values:", values)

    # Filter rows by item name
    matching_rows = [row for row in values if row and row[0].strip().lower() == item_name.strip().lower()]
    
    # Debug: Print matching rows
    print("Matching rows:", matching_rows)
    
    return matching_rows

@app.route('/', methods=['GET', 'POST'])
def index():
    search_results = None
    if request.method == 'POST':
        item_name = request.form.get('item')
        # Debug: Print the item name being searched
        print("Searching for item:", item_name)
        search_results = search_item_in_sheet(SPREADSHEET_ID, item_name)
    
    return render_template('search.html', search_results=search_results)

if __name__ == '__main__':
    app.run(debug=True, port=5001)
