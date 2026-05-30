import os
import gspread
from dotenv import load_dotenv

load_dotenv()

def sync_excel_to_gsheet(excel_file):
    import pandas as pd
    from gspread_dataframe import set_with_dataframe

    sheet_name = os.getenv("GOOGLE_SHEET_NAME")
    creds_file = os.getenv("GOOGLE_CREDENTIALS_JSON") or os.getenv("GOOGLE_CREDENTIALS_FILE")

    if not sheet_name or not creds_file:
        print("Missing Google Sheets config in .env")
        return

    client = gspread.service_account(filename=creds_file)

    print(f"Syncing Excel to Google Sheet: {sheet_name}")
    sheet = client.open(sheet_name)

    xls = pd.read_excel(excel_file, sheet_name=None)

    for tab, df in xls.items():
        print(f"Updating tab: {tab}")
        try:
            ws = sheet.worksheet(tab)
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet(title=tab, rows=1000, cols=26)
        ws.clear()
        set_with_dataframe(ws, df)


def sync_wishlist_to_gsheet(items):
    sheet_name = os.getenv("GOOGLE_SHEET_NAME")
    creds_file = os.getenv("GOOGLE_CREDENTIALS_JSON") or os.getenv("GOOGLE_CREDENTIALS_FILE")

    if not sheet_name or not creds_file:
        raise ValueError("Missing GOOGLE_SHEET_NAME or GOOGLE_CREDENTIALS_JSON in .env")

    client = gspread.service_account(filename=creds_file)
    sheet  = client.open(sheet_name)

    try:
        ws = sheet.worksheet("Wishlist")
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title="Wishlist", rows=1000, cols=10)

    rows = [["Title", "Category", "Release Year", "Notes", "Added"]]
    rows += [[i["title"], i["type"], i["release"], i["notes"], i["created_at"]] for i in items]

    ws.clear()
    ws.update(rows, value_input_option="RAW")
    return len(items)

