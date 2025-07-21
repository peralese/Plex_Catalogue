import os
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from plexapi.server import PlexServer

from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

# ──────────────────────────────────────────────────────────────────────────────
# 1. Environment
# ──────────────────────────────────────────────────────────────────────────────
load_dotenv()
BASEURL = os.getenv("PLEX_BASEURL")
TOKEN = os.getenv("PLEX_TOKEN")

IGNORE_LIBRARIES = [
    lib.strip() for lib in os.getenv("IGNORE_LIBRARIES", "").split(",") if lib.strip()
]

if not BASEURL or not TOKEN:
    raise ValueError("PLEX_BASEURL and PLEX_TOKEN must be set in .env")

BACKUP_TAGS = ["dvd", "blue-ray", "iso", "ripped"]  # canonical lowercase tags

# ──────────────────────────────────────────────────────────────────────────────
# 2. Connect to Plex
# ──────────────────────────────────────────────────────────────────────────────
plex = PlexServer(BASEURL, TOKEN)

# ──────────────────────────────────────────────────────────────────────────────
# 3. Output workbook
# ──────────────────────────────────────────────────────────────────────────────
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
out_dir = os.path.join("output", timestamp)
os.makedirs(out_dir, exist_ok=True)

excel_path = os.path.join(out_dir, "plex_media_catalog.xlsx")
writer = pd.ExcelWriter(excel_path, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Helper: autosize columns
# ──────────────────────────────────────────────────────────────────────────────
def autosize(ws, pad=2, max_width=80):
    """Auto‑fit each column in an openpyxl worksheet."""
    # for col_cells in ws.columns:
    #     max_len = max((len(str(c.value)) for c in col_cells if c.value), default=0)
    #     ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + pad, max_width)
    for col_cells in ws.columns:
        # Skip empty columns
        col_letter = None
        for cell in col_cells:
            if not isinstance(cell, type(ws.cell(row=1, column=1))):  # regular Cell
                continue
            col_letter = cell.column_letter
            break

        if col_letter:
            max_len = max((len(str(c.value)) for c in col_cells if c.value), default=0)
            ws.column_dimensions[col_letter].width = min(max_len + pad, max_width)

# ──────────────────────────────────────────────────────────────────────────────
# Helper: collections → tag list
# ──────────────────────────────────────────────────────────────────────────────
def get_collection_tags(item):
    try:
        return [c.tag.lower() for c in item.collections]
    except Exception:
        return []

# ──────────────────────────────────────────────────────────────────────────────
# Helper: classify backup (returns set of backup tags found & bool backed_up)
# ──────────────────────────────────────────────────────────────────────────────
def detect_backup(tags, file_path):
    """Return set({'dvd','iso',...}) and backed_up flag."""
    found = set(t for t in BACKUP_TAGS if t in tags)

    # fallback by file path
    fp = file_path.lower()
    if ".iso" in fp:
        found.add("iso")
    if "dvd" in fp or ".vob" in fp:
        found.add("dvd")

    return found, bool(found)

# ──────────────────────────────────────────────────────────────────────────────
# 4. Process libraries
# ──────────────────────────────────────────────────────────────────────────────
dashboard = []  # rows for dashboard
movie_sheets = []
tv_rows = []

for section in plex.library.sections():
    if section.title in IGNORE_LIBRARIES:
        print(f"Skipping library: {section.title}")
        continue

    if section.type == "movie":
        print("Movies:", section.title)
        rows = []
        stats = {"dvd": 0, "blue-ray": 0, "iso": 0, "ripped": 0, "total": 0}

        for movie in section.all():
            try:
                tags = get_collection_tags(movie)
                file_path = movie.media[0].parts[0].file
                found, backed = detect_backup(tags, file_path)

                # stats
                stats["total"] += 1
                for tag in found:
                    stats[tag] += 1

                # store row
                rows.append(
                    [
                        movie.title,
                        "Yes" if backed else "No",
                        ", ".join(tag.upper() if tag != "blue-ray" else "Blue-ray" for tag in found),
                        file_path,
                    ]
                )
            except Exception as exc:
                print("  ↳ skipped movie:", exc)

        # convert & append totals row
        df = pd.DataFrame(rows, columns=["Title", "Backup", "Type", "Path"])
        df["% Backed Up"] = ""
        df = pd.concat(
            [
                df,
                pd.DataFrame(
                    [
                        [
                            "Total",
                            stats["dvd"] + stats["blue-ray"] + stats["iso"] + stats["ripped"],
                            "",
                            "",
                            "",
                        ]
                    ],
                    columns=["Title", "Backup", "Type", "Path", "% Backed Up"],
                ),
            ],
            ignore_index=True,
        )

        movie_sheets.append((section.title[:31], df))

        # add to dashboard
        total = stats["total"]
        dvd, br, iso, rip = stats["dvd"], stats["blue-ray"], stats["iso"], stats["ripped"]
        pct = round((dvd + br + iso + rip) / total * 100, 1) if total else 0
        dashboard.append([section.title, total, dvd, br, iso, rip, pct])

    elif section.type == "show":
        print("TV:", section.title)
        for show in section.all():
            show_tags = get_collection_tags(show)
            try:
                for ep in show.episodes():
                    ep_tags = get_collection_tags(ep)
                    tags = ep_tags or show_tags
                    file_path = ep.media[0].parts[0].file
                    found, backed = detect_backup(tags, file_path)

                    tv_rows.append(
                        [
                            show.title,
                            ep.seasonNumber,
                            ep.index,
                            ep.title,
                            "Yes" if backed else "No",
                            ", ".join(found).upper(),
                            file_path,
                        ]
                    )
            except Exception as exc:
                print("  ↳ skipped episode:", exc)

# ──────────────────────────────────────────────────────────────────────────────
# 5. Build Dashboard sheet
# ──────────────────────────────────────────────────────────────────────────────
dash_cols = [
    "Category",
    "Movie Count",
    "DVD",
    "Blue-ray",
    "ISO",
    "Ripped",
    "% Backed Up",
]
dash_df = pd.DataFrame(dashboard, columns=dash_cols)  # pct added next
dash_df["% Backed Up"] = dash_df.apply(
    lambda r: round((r["DVD"] + r["Blue-ray"] + r["ISO"] + r["Ripped"]) / r["Movie Count"] * 100, 1)
    if r["Movie Count"]
    else 0,
    axis=1,
)

# total row
tot_row = pd.DataFrame(
    [
        [
            "Total",
            dash_df["Movie Count"].sum(),
            dash_df["DVD"].sum(),
            dash_df["Blue-ray"].sum(),
            dash_df["ISO"].sum(),
            dash_df["Ripped"].sum(),
            round(
                (dash_df["DVD"].sum() + dash_df["Blue-ray"].sum() + dash_df["ISO"].sum() + dash_df["Ripped"].sum())
                / dash_df["Movie Count"].sum()
                * 100,
                1,
            )
            if dash_df["Movie Count"].sum()
            else 0,
        ]
    ],
    columns=dash_cols,
)
dash_df = pd.concat([dash_df, tot_row], ignore_index=True)

# write starting at row 3 to leave space for 2‑row header
dash_df.to_excel(
    writer, sheet_name="Dashboard", index=False, startrow=2, header=False
)
ws_dash = writer.sheets["Dashboard"]

# write custom 2‑row header
headers_top = ["Category", "Movie Count", "Backup Type", "", "", "", "% Backed Up"]
headers_mid = ["", "", "DVD", "Blue-ray", "ISO", "Ripped", ""]
# ws_dash.append(headers_top)
# ws_dash.append(headers_mid)

# Write headers manually
for i, header_row in enumerate([headers_top, headers_mid], start=1):
    for j, val in enumerate(header_row, start=1):
        ws_dash.cell(row=i, column=j, value=val)

# Merge 'Backup Type' header
ws_dash.merge_cells(start_row=1, start_column=3, end_row=1, end_column=6)

# Bold headers
from openpyxl.styles import Font
for row in ws_dash.iter_rows(min_row=1, max_row=2):
    for cell in row:
        cell.font = Font(bold=True)

# move data (dash_df) two rows down already handled via startrow
# merge cells for "Backup Type"
ws_dash.merge_cells(start_row=1, start_column=3, end_row=1, end_column=6)
ws_dash["C1"].alignment = Alignment(horizontal="center")

# bold header rows
for row in ws_dash["1:2"]:
    for cell in row:
        cell.font = Font(bold=True)

autosize(ws_dash)

# Add pie chart below table (start around row 10)
chart = PieChart()
chart.title = "Backup Type Distribution (Total)"
chart.height = 7  # default = 7.5
chart.width = 10  # default = 15

# Find the last row with data (the "Total" row)
last_row = ws_dash.max_row
# Step up one row if the chart is picking up wrong cells (because the last row might be empty or chart follows totals)
while not ws_dash.cell(row=last_row, column=1).value or str(ws_dash.cell(row=last_row, column=1).value).strip() == "":
    last_row -= 1

# Force numeric cell values in Total row for DVD–Ripped (columns C–F)
for col_idx in range(3, 7):  # Columns C to F
    cell = ws_dash.cell(row=last_row, column=col_idx)
    try:
        cell.value = int(cell.value)
    except (ValueError, TypeError):
        cell.value = 0


# Reference backup type values (DVD to Ripped) and labels
data = Reference(ws_dash, min_col=3, max_col=6, min_row=last_row, max_row=last_row)
labels = Reference(ws_dash, min_col=3, max_col=6, min_row=2)

chart.add_data(data, titles_from_data=False)
chart.set_categories(labels)
chart.dataLabels = DataLabelList()
chart.dataLabels.showVal = True
chart.dataLabels.showPercent = True

# Place chart starting a few rows below the table
ws_dash.add_chart(chart, "B{}".format(last_row + 2))


# ──────────────────────────────────────────────────────────────────────────────
# 6. Movie sheets
# ──────────────────────────────────────────────────────────────────────────────
for sheet_name, df in movie_sheets:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    autosize(writer.sheets[sheet_name])

# ──────────────────────────────────────────────────────────────────────────────
# 7. TV sheet
# ──────────────────────────────────────────────────────────────────────────────
if tv_rows:
    tv_df = pd.DataFrame(
        tv_rows,
        columns=[
            "Show Title",
            "Season",
            "Episode",
            "Episode Title",
            "Backup",
            "Type",
            "Path",
        ],
    )
    tv_df.to_excel(writer, sheet_name="TV_Shows", index=False)
    autosize(writer.sheets["TV_Shows"])

# ──────────────────────────────────────────────────────────────────────────────
# 8. Wishlist
# ──────────────────────────────────────────────────────────────────────────────
pd.DataFrame(columns=["Title", "Notes", "Desired Format"]).to_excel(
    writer, sheet_name="Wishlist", index=False
)
autosize(writer.sheets["Wishlist"])

# ──────────────────────────────────────────────────────────────────────────────
writer.close()
print("✅ Excel saved ➜", excel_path)

