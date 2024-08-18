import os
import pickle
from flask import Flask, render_template, request
import google.auth.transport.requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import datetime
import config

app = Flask(__name__)

def get_creds(scopes):
    creds = None
    token_path = os.path.join(os.getcwd(), 'token.pickle')
    
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                config.CLIENT_SECRET_FILE, scopes)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def get_sheet_id(service, spreadsheet_id, sheet_name):
    """Get the numeric sheet ID for a given sheet name."""
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in spreadsheet['sheets']:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    raise Exception(f"Sheet with name '{sheet_name}' not found.")

def find_next_empty_row(service, spreadsheet_id, sheet_name):
    """Finds the next empty row for a given sheet."""
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:A"
    ).execute()
    values = result.get('values', [])
    return len(values) + 1  # Return the next row number

def append_to_sheet(item, location, quantity):
    creds = get_creds(config.SCOPES_EDIT)
    service = build('sheets', 'v4', credentials=creds)
    
    # Find the next empty row in the sheet
    next_row = find_next_empty_row(service, config.SPREADSHEET_ID, config.RANGE_NAME.split('!')[0])
    
    # Generate the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Data to be inserted
    data = [item, location, quantity, timestamp]
    
    # Specify the range to insert the data
    range_name = f'{config.RANGE_NAME.split("!")[0]}!A{next_row}:D{next_row}'
    
    body = {'values': [data]}
    service.spreadsheets().values().update(
        spreadsheetId=config.SPREADSHEET_ID,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()

def clear_location(location):
    """Clears all entries with the specified location."""
    creds = get_creds(config.SCOPES_EDIT)
    service = build('sheets', 'v4', credentials=creds)

    sheet_name = config.RANGE_NAME.split('!')[0]
    sheet_id = get_sheet_id(service, config.SPREADSHEET_ID, sheet_name)
    
    range_name = config.RANGE_NAME
    
    # Fetch all rows from the sheet
    result = service.spreadsheets().values().get(
        spreadsheetId=config.SPREADSHEET_ID,
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
        spreadsheetId=config.SPREADSHEET_ID,
        body=body
    ).execute()

def search_item_in_sheet(item_name):
    """Searches for an item in the Google Sheet and returns matching rows."""
    creds = get_creds(config.SCOPES_READONLY)
    service = build('sheets', 'v4', credentials=creds)
    
    range_name = config.RANGE_NAME
    result = service.spreadsheets().values().get(
        spreadsheetId=config.SPREADSHEET_ID,
        range=range_name
    ).execute()
    values = result.get('values', [])

    # Filter rows by item name
    matching_rows = [row for row in values if row and row[0].strip().lower() == item_name.strip().lower()]
    
    return matching_rows

@app.route('/', methods=['GET', 'POST'])
def index():
    search_results = None
    item_name = None  # Start with item_name as None
    
    if request.method == 'POST':
        if 'append' in request.form:
            item = request.form.get('item')
            location = request.form.get('location')
            quantity = request.form.get('quantity')
            append_to_sheet(item, location, quantity)
        elif 'clear' in request.form:
            location = request.form.get('location')
            clear_location(location)
        elif 'search' in request.form:
            item_name = request.form.get('item').strip()  # Strip any extra whitespace
            if item_name:  # Only perform search if item_name is not empty
                search_results = search_item_in_sheet(item_name)
    
    return render_template('index.html', 
                           search_results=search_results, 
                           item_name=item_name, 
                           spreadsheet_id=config.SPREADSHEET_ID)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
