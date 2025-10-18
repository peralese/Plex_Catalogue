
import os
from datetime import datetime
import shutil

import pandas as pd
from dotenv import load_dotenv
from openpyxl.styles import Font, Alignment
from openpyxl.chart.label import DataLabelList
from openpyxl.cell.cell import MergedCell 
from plexapi.server import PlexServer
from openpyxl.cell.cell import MergedCell           
from openpyxl.chart import BarChart
from openpyxl.chart.reference import Reference
from modules.google_sync import sync_excel_to_gsheet
from modules.movie_wishlist_sync import write_wishlist_to_excel