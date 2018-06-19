import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('gspreadcreds.json', scope)
client = gspread.authorize(creds)

google_sheet = client.open("GSpread-Sample").sheet1

file_name = "sevarthis.xlsx"
sevarthis_df = pd.read_excel(file_name, sheet_name="Unique_list")

sheet_row_index = 1
for i, row in sevarthis_df.iterrows():
    row = row.to_dict()

    if i == 0:
        headers = list(row.keys()) + ["Month", "Seva Attendance Days Per Month", "Total Session Per Month", "Remarks"]
        google_sheet.insert_row(headers, sheet_row_index)

    sheet_row_index += 1
    google_sheet.insert_row(list(row.values()), sheet_row_index) 

print("Data saved to Google's Spread Sheet")