web: python app.py
heroku ps:scale web=1
heroku config:set SPREADSHEET_ID = '1eAZt-gjE_BpciYDZENUwZ5MUI9L7vSjZ16a7yL5vuo8'
heroku config:set RANGE_NAME = 'Sheet1!A:D'
heroku config:set CLIENT_SECRET_FILE = r'main\client_secret.json'
heroku config:set SCOPES_READONLY = ['https://www.googleapis.com/auth/spreadsheets.readonly']
heroku config:set SCOPES_EDIT = ['https://www.googleapis.com/auth/spreadsheets']
