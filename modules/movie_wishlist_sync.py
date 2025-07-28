import os
import pandas as pd
import gspread
from dotenv import load_dotenv

load_dotenv()
CREDENTIALS_PATH = os.getenv("GOOGLE_CREDENTIALS_JSON", "google_credentials.json")
WISHLIST_SHEET_NAME = os.getenv("MOVIE_WISHLIST_SHEET", "DVD Wish List")

def load_movie_wishlist(sheet_name=WISHLIST_SHEET_NAME):
    gc = gspread.service_account(filename=CREDENTIALS_PATH)
    sheet = gc.open(sheet_name).sheet1
    df = pd.DataFrame(sheet.get_all_records())
    return df

def save_movie_wishlist(sheet_name, df):
    gc = gspread.service_account(filename=CREDENTIALS_PATH)
    sheet = gc.open(sheet_name).sheet1
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())


def write_wishlist_to_excel(writer, sheet_name="Wishlist"):
    gc = gspread.service_account(filename=CREDENTIALS_PATH)
    sheet = gc.open(WISHLIST_SHEET_NAME).sheet1
    df = pd.DataFrame(sheet.get_all_records())
    df.to_excel(writer, sheet_name=sheet_name, index=False)

