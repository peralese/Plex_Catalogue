import os
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from plexapi.server import PlexServer

# ---------------------------------------------------------------------------
# 1. Configuration
# ---------------------------------------------------------------------------
load_dotenv()

BASEURL = os.getenv("PLEX_BASEURL")
TOKEN = os.getenv("PLEX_TOKEN")
IGNORE_LIBRARIES = [
    lib.strip() for lib in os.getenv("IGNORE_LIBRARIES", "").split(",") if lib.strip()
]

if not BASEURL or not TOKEN:
    raise ValueError("PLEX_BASEURL and PLEX_TOKEN must be set in .env")

BACKUP_KEYWORDS = ["backup", "iso", "dvd", "blue-ray"]

# ---------------------------------------------------------------------------
# 2. Connect to Plex
# ---------------------------------------------------------------------------
plex = PlexServer(BASEURL, TOKEN)

# ---------------------------------------------------------------------------
# 3. Prepare output paths
# ---------------------------------------------------------------------------
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_dir = os.path.join("output", timestamp)
os.makedirs(output_dir, exist_ok=True)
excel_path = os.path.join(output_dir, "plex_media_catalog.xlsx")

writer = pd.ExcelWriter(excel_path, engine="openpyxl")

# ---------------------------------------------------------------------------
# 4. Helpers
# ---------------------------------------------------------------------------
def extract_tags_from_collections(item):
    """
    Return a lowercase list of collection names for a Movie, Show, or Episode.
    plexapi represents collections as list[str] (titles) in .collections.
    """
    try:
        return [c.tag.lower() for c in item.collections]
    except Exception:
        return []


def classify_backup(tags, file_path):
    """
    Decide Backup (Yes/No) and media_type (ISO/DVD/Blue‑ray/'') using:
    • Primary: collection tags
    • Fallback: file‑path inspection
    """
    matched_types = []
    # 1) collections
    for key in ["iso", "dvd", "blue-ray"]:
        if key in tags:
            matched_types.append(key.upper() if key != "blue-ray" else "Blue-ray")
    
    if "backup" in tags or matched_types:
        backup = "Yes"
        media_type = ", ".join(matched_types)
        return backup, media_type

    # if any(t in tags for t in BACKUP_KEYWORDS):
    #     backup = "Yes"
    #     if "iso" in tags:
    #         media_type = "ISO"
    #     elif "dvd" in tags:
    #         media_type = "DVD"
    #     elif "blue-ray" in tags:
    #         media_type = "Blue-ray"
    #     else:
    #         media_type = ""
    #     return backup, media_type

    # 2) path fallback
    fp = file_path.lower()
    if ".iso" in fp:
        return "Yes", "ISO"
    if "dvd" in fp or ".vob" in fp:
        return "Yes", "DVD"
    return "No", ""
# ---------------------------------------------------------------------------
# 5. Iterate Plex libraries
# ---------------------------------------------------------------------------
dashboard_rows = []
movie_sheets = []
tv_rows = []

for section in plex.library.sections():
    if section.title in IGNORE_LIBRARIES:
        print(f"Skipping ignored library: {section.title}")
        continue

    if section.type == "movie":
        print(f"Processing movie library: {section.title}")
        rows = []

        for movie in section.all():

            try:
                tags = extract_tags_from_collections(movie)
                file_path = movie.media[0].parts[0].file
                backup, media_type = classify_backup(tags, file_path)
                rows.append([movie.title, backup, media_type, file_path])
            except Exception as exc:
                print(f"  ⤷ Skipped movie due to error: {exc}")

        df = pd.DataFrame(rows, columns=["Title", "Backup", "Type", "Path"])

        # totals row
        total = len(df)
        backed_up = len(df[df["Backup"] == "Yes"])
        pct = round((backed_up / total) * 100, 1) if total else 0
        df["% Backed Up"] = ""
        total_row = pd.DataFrame(
            [["Total", backed_up, "", "", f"{pct}%"]],
            columns=["Title", "Backup", "Type", "Path", "% Backed Up"],
        )
        df = pd.concat([df, total_row], ignore_index=True)

        movie_sheets.append((section.title[:31], df))
        dashboard_rows.append([section.title, total, backed_up, pct])

    elif section.type == "show":
        for show in section.all():
            show_tags = extract_tags_from_collections(show)

            try:
                for ep in show.episodes():
                    ep_tags = extract_tags_from_collections(ep)
                    tags = ep_tags or show_tags
                    file_path = ep.media[0].parts[0].file
                    backup, media_type = classify_backup(tags, file_path)
                    tv_rows.append(
                        [
                            show.title,
                            ep.seasonNumber,
                            ep.index,
                            ep.title,
                            backup,
                            media_type,
                            file_path,
                        ]
                    )
            except Exception as exc:
                print(f"  ⤷ Skipped show/episode due to error: {exc}")

# ---------------------------------------------------------------------------
# 6. Write Dashboard first
# ---------------------------------------------------------------------------
dash_df = pd.DataFrame(
    dashboard_rows, columns=["Category", "Total Movies", "With Backup", "% Backed Up"]
)
dash_total = pd.DataFrame(
    [
        [
            "Total",
            dash_df["Total Movies"].sum(),
            dash_df["With Backup"].sum(),
            round(
                dash_df["With Backup"].sum() / dash_df["Total Movies"].sum() * 100, 1
            )
            if dash_df["Total Movies"].sum()
            else 0,
        ]
    ],
    columns=dash_df.columns,
)
dash_df = pd.concat([dash_df, dash_total], ignore_index=True)
dash_df.to_excel(writer, sheet_name="Dashboard", index=False)

# ---------------------------------------------------------------------------
# 7. Movie sheets
# ---------------------------------------------------------------------------
for sheet_name, df in movie_sheets:
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# ---------------------------------------------------------------------------
# 8. TV Shows sheet
# ---------------------------------------------------------------------------
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
    total = len(tv_df)
    backed_up = len(tv_df[tv_df["Backup"] == "Yes"])
    pct = round((backed_up / total) * 100, 1) if total else 0
    tv_df["% Backed Up"] = ""
    tv_total = pd.DataFrame(
        [["Total", "", "", "", backed_up, "", "", f"{pct}%"]],
        columns=[
            "Show Title",
            "Season",
            "Episode",
            "Episode Title",
            "Backup",
            "Type",
            "Path",
            "% Backed Up",
        ],
    )
    tv_df = pd.concat([tv_df, tv_total], ignore_index=True)
    tv_df.to_excel(writer, sheet_name="TV_Shows", index=False)

# ---------------------------------------------------------------------------
# 9. Wishlist sheet
# ---------------------------------------------------------------------------
pd.DataFrame(columns=["Title", "Notes", "Desired Format"]).to_excel(
    writer, sheet_name="Wishlist", index=False
)

writer.close()
print(f"✅ Excel file saved →  {excel_path}")

