# modules/movie_wishlist_sync.py
import os
import pandas as pd
import gspread
from dotenv import load_dotenv

load_dotenv()
CREDENTIALS_PATH = os.getenv("GOOGLE_CREDENTIALS_JSON", "google_credentials.json")
WISHLIST_SHEET_NAME = os.getenv("MOVIE_WISHLIST_SHEET", "DVD Wish List")

def write_wishlist_to_excel(writer, sheet_name="Wishlist"):
    gc = gspread.service_account(filename=CREDENTIALS_PATH)
    sheet = gc.open(WISHLIST_SHEET_NAME).sheet1
    df = pd.DataFrame(sheet.get_all_records())
    df.to_excel(writer, sheet_name=sheet_name, index=False)

