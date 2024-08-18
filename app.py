import os
import pickle
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for
import google.auth.transport.requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import config

app = Flask(__name__)

def get_creds(scopes):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                config.CLIENT_SECRET_FILE, scopes)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def get_sheet_id(service, spreadsheet_id, sheet_name):
    """Retrieves the numeric sheet ID for the given sheet name."""
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    
    return None  # Sheet not found

def create_sheet_if_not_exists(service, spreadsheet_id, sheet_name):
    """Creates a new sheet if it doesn't exist."""
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheet_titles = [sheet['properties']['title'] for sheet in sheets]

    if sheet_name not in sheet_titles:
        requests = [{
            "addSheet": {
                "properties": {
                    "title": sheet_name
                }
            }
        }]
        body = {"requests": requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

def find_next_empty_row(service, spreadsheet_id, sheet_name):
    """Finds the next empty row for a given sheet."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:A"
    ).execute()
    values = result.get('values', [])
    return len(values) + 1  # Return the next row number

def append_to_sheet(spreadsheet_id, item, location, quantity):
    creds = get_creds(config.SCOPES_EDIT)
    service = build('sheets', 'v4', credentials=creds)
    
    # Sheet name where data will be stored
    sheet_name = "Sheet1"
    
    # Ensure the sheet exists
    create_sheet_if_not_exists(service, spreadsheet_id, sheet_name)
    
    # Find the next empty row in the sheet
    next_row = find_next_empty_row(service, spreadsheet_id, sheet_name)
    
    # Generate the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Data to be inserted
    data = [item, location, quantity, timestamp]
    
    # Specify the range to insert the data
    range_name = f'{sheet_name}!A{next_row}:D{next_row}'
    
    body = {'values': [data]}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()

def clear_location(spreadsheet_id, location):
    """Clears all entries with the specified location."""
    creds = get_creds(config.SCOPES_EDIT)
    service = build('sheets', 'v4', credentials=creds)
    
    sheet_name = "Sheet1"
    sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name)
    if sheet_id is None:
        raise ValueError(f"Sheet {sheet_name} not found.")
    
    range_name = f"{sheet_name}!A:D"
    
    # Fetch all rows from the sheet
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    values = result.get('values', [])
    
    # Identify rows that match the location
    rows_to_clear = []
    for idx, row in enumerate(values):
        if len(row) > 1 and row[1].strip().lower() == location.strip().lower():
            rows_to_clear.append(idx + 1)  # Add 1 because sheet row numbers are 1-based
    
    if not rows_to_clear:
        return  # No matching rows to clear
    
    # Prepare batch update to clear rows
    requests = []
    for row_index in reversed(rows_to_clear):  # Reverse to avoid messing up row indices
        requests.append({
            "deleteRange": {
                "range": {
                    "sheetId": sheet_id,  # Use the numeric sheet ID
                    "startRowIndex": row_index - 1,  # Convert to 0-based index
                    "endRowIndex": row_index,  # Clear only this row
                    "startColumnIndex": 0,
                    "endColumnIndex": 4  # A:D columns
                },
                "shiftDimension": "ROWS"
            }
        })
    
    body = {"requests": requests}
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()

@app.route('/', methods=['GET', 'POST'])
def index():
    # Create a new sheet if it doesn't exist
    if not os.path.exists('spreadsheet_id.txt'):
        creds = get_creds(config.SCOPES_EDIT)
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet = service.spreadsheets().create(body={
            'properties': {'title': 'Inventory Tracker'}
        }).execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')
        with open('spreadsheet_id.txt', 'w') as f:
            f.write(spreadsheet_id)
        
        # Create the initial sheet
        create_sheet_if_not_exists(service, spreadsheet_id, "Sheet1")
    else:
        with open('spreadsheet_id.txt', 'r') as f:
            spreadsheet_id = f.read().strip()
    
    if request.method == 'POST':
        # Handle appending data
        if 'append' in request.form:
            item = request.form.get('item')
            location = request.form.get('location')
            quantity = request.form.get('quantity')
            append_to_sheet(spreadsheet_id, item, location, quantity)
        # Handle clearing location
        elif 'clear' in request.form:
            location = request.form.get('location')
            clear_location(spreadsheet_id, location)
        
        return redirect(url_for('index'))

    # Create the Google Sheets URL
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    
    return render_template('index.html', sheet_url=sheet_url)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
