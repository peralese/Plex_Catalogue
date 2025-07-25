import os
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

load_dotenv()

def sync_excel_to_gsheet(excel_file):
    sheet_name = os.getenv("GOOGLE_SHEET_NAME")
    creds_file = os.getenv("GOOGLE_CREDENTIALS_FILE")

    if not sheet_name or not creds_file:
        print("Missing Google Sheets config in .env")
        return

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
    client = gspread.authorize(creds)

    print(f"🔄 Syncing Excel to Google Sheet: {sheet_name}")
    sheet = client.open(sheet_name)

    # Read all tabs from Excel
    xls = pd.read_excel(excel_file, sheet_name=None)

    for tab, df in xls.items():
        print(f"→ Updating tab: {tab}")
        try:
            ws = sheet.worksheet(tab)
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet(title=tab, rows=1000, cols=26)
        ws.clear()
        set_with_dataframe(ws, df)
