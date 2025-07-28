import os
import gspread
from dotenv import load_dotenv

load_dotenv()
SHEET_NAME = os.getenv("MOVIE_WISHLIST_SHEET", "DVD Wish List")
CREDS_FILE = os.getenv("GOOGLE_CREDENTIALS_JSON", "google_credentials.json")

gc = gspread.service_account(filename=CREDS_FILE)
sheet = gc.open(SHEET_NAME).sheet1

def get_wishlist():
    return sheet.get_all_records()

def add_wishlist_item(title, notes, fmt):
    sheet.append_row([title, notes, fmt])

def delete_item(index):
    # +2 because header is row 1 and index is 0-based
    sheet.delete_rows(index + 2)

def update_item(index, title, notes, fmt):
    row = index + 2
    sheet.update(f"A{row}:C{row}", [[title, notes, fmt]])