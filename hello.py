"""
Shows basic usage of the Sheets API. Prints values from a Google Spreadsheet.
"""
from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# Setup the Sheets API
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
store = file.Storage('credentials.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = build('sheets', 'v4', http=creds.authorize(Http()))

# Call the Sheets API
SPREADSHEET_ID = '1_iqPh4nnRgIRSnE5SzdaXrN7ZOfD-y_na2-bAWMpDyQ'
RANGE_NAME = 'Class Data!A:E'
result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
                                             range=RANGE_NAME).execute()
values = result.get('values', [])
count=0
if not values:
    print('No data found.')
else:
    #print('Name, Major:')
    for row in values:
        # Print columns A and E, which correspond to indices 0 and 4.
        #print('%s, %s, %s, %s, %s' % (row[0], row[1], row[2], row[3], row[4]))
        if row[4]=="E":
            count += 1
    print("count of E = ", count)
dest_spread = '1xJC7GewEHCS6oqJrbcJ8i5NDyQ6MZQqrI7JV-GgX-pw'
dest_range = 'Sheet1!A1'
value_input_option = 'RAW'
request = service.spreadsheets().values().update(spreadsheetId=dest_spread,range=dest_range, valueInputOption='RAW',body=values)
response = request.execute()
print(response)