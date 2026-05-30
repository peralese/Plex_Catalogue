# Plex Catalog Exporter

Exports your Plex movie and TV libraries into an organized Excel workbook, with dashboards, backup tracking, and a Google Sheets integration for sharing and wishlist sync.

---

## Features

- Exports one Excel tab per movie library (e.g., Movies, Classics, Anime)
- Generates a TV Shows summary sheet and a TV Dashboard with pie chart
- Creates a Dashboard tab summarizing movie backup stats by type
- Includes a bar chart of movie backup types by library
- Automatically uploads the Excel workbook to Google Sheets (configurable)
- Pulls a movie wishlist from an external Google Sheet
- Web UI to view/edit the DVD Wish List
- Saves exports in timestamped folders (e.g., `output/2025-07-21_13-00-22/`)

---

## Output Structure

Each Excel export includes:

| Sheet Name                | Description                                   |
|---------------------------|-----------------------------------------------|
| `Dashboard`               | Backup summary per movie library + chart      |
| `TV_Dashboard`            | TV shows summary + pie chart                  |
| `Movies`, `Classics`, …   | One tab per Plex movie library                |
| `TV_Shows`                | Combined list of all TV shows                 |
| `Wishlist`                | Pulled live from external Google Sheet        |

---

## Requirements

- Python 3.9+
- Plex Media Server + token
- Google Cloud service account with access to your Sheets

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Configuration (.env)

Required:

```env
PLEX_BASEURL=http://localhost:32400
PLEX_TOKEN=your_token_here

# Target Google Sheet (destination for export)
GOOGLE_SHEET_NAME=Plex Movies
# Service account credentials JSON
GOOGLE_CREDENTIALS_JSON=google_credentials.json

# Source Google Sheet for Wishlist
MOVIE_WISHLIST_SHEET=DVD Wish List
```

Optional:

```env
# Skip export to Google Sheets if false/0/no/off
SYNC_TO_GOOGLE=true
# Comma-separated list of Plex libraries to skip
IGNORE_LIBRARIES=Music Videos, Kids Movies
```

Share both Google Sheets with your service account email:
- Target sheet to receive the export: `GOOGLE_SHEET_NAME`
- Wishlist source sheet: `MOVIE_WISHLIST_SHEET`

---

## How to Use

Run the exporter:

```bash
python plex_catalog_exporter.py
```

Launch the web UI (optional):

```bash
python -m app.app
```

Open http://localhost:5000 to view/edit the wishlist.

---

## Sync Behavior

- Overwrites each tab in the Google Sheet that matches the Excel sheets
- Extra tabs (e.g., `Notes`) in your Google Sheet are left untouched
- The wishlist is pulled live from Google Sheets at runtime
- Toggle syncing via `SYNC_TO_GOOGLE` (defaults to `true`). When disabled, the Excel is generated but not uploaded.
- Only cell values are synced to Google Sheets (formatting and charts remain in the local Excel file).

---

## Backup Tags Logic

Backup types (`DVD`, `ISO`, `Blu-ray`, `Ripped`) are pulled from the Labels field in Plex metadata.

- Canonical tag names: `dvd`, `blu-ray`, `iso`, `ripped`
- Recognized Blu-ray aliases: `blue-ray`, `bluray` (treated as `blu-ray`)
- Fallback detection: file paths containing `.iso` imply ISO; `dvd` or `.vob` imply DVD
- Exclude whole libraries with `IGNORE_LIBRARIES` (comma‑separated)

Add these labels to your Plex movies or episodes to track backup types. Multiple labels per item are supported.

---

## Roadmap

- [x] Replace local Wishlist tab with live data from Google Sheets
- [x] Automatically sync final Excel output to Google Sheets
- [x] Show bar chart of movie backup types
- [x] Add pie chart of TV episode backup coverage
- [x] Add TV Dashboard tab
- [x] Switch from Collections to Labels for backup tagging
- [x] Add web UI for viewing/editing the wish list
- [x] Auto-cleanup old timestamped folders after successful upload
- [x] Export the wishlist to Google Sheets

---

## License

MIT License

You are free to use, modify, and distribute this tool with attribution.

---

## Author

**Erick Perales** - IT Architect, Cloud Migration Specialist  
https://github.com/peralese
