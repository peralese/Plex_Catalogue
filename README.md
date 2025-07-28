# Plex Catalog Exporter

This Python script exports your Plex movie and TV libraries into an organized Excel spreadsheet, including dashboards, backup tracking, and a wish list integration from Google Sheets.

---

## 📦 Features

- 🔍 Exports **one Excel tab per movie library** (e.g., Movies, Classics, Anime)
- 📺 Generates a **TV Shows summary** sheet and a **TV Dashboard** with pie chart
- 📊 Creates a **Dashboard** tab summarizing movie backup stats by type
- 📁 Includes a **bar chart** of movie backup types by library
- ☁️ **Automatically uploads the Excel workbook to Google Sheets**
- 🌐 **Pulls a movie wishlist** from an external Google Sheet
- 🖥️ **Includes a web-based UI** to edit the DVD Wish List interactively
- 📁 Saves exports in timestamped folders (e.g., `output/catalog_2025-07-21_130022/`)

---

## 📂 Output Structure

Each Excel export includes:

| Sheet Name         | Description                                      |
|--------------------|--------------------------------------------------|
| `Dashboard`        | Backup summary per movie library + chart         |
| `TV_Dashboard`     | TV shows summary + pie chart                     |
| `Movies`, `Classics`, etc. | One tab per Plex movie library         |
| `TV_Shows`         | Combined list of all TV shows                    |
| `Wishlist`         | Pulled live from external Google Sheet           |

---

## ✅ Requirements

- Python 3.9+
- A Plex Media Server
- A service account with access to Google Sheets
- `.env` file with the following:

```env
PLEX_BASEURL=http://localhost:32400
PLEX_TOKEN=your_token_here
IGNORE_LIBRARIES=Music Videos

GOOGLE_CREDENTIALS_JSON=google_credentials.json
GOOGLE_SHEET_NAME=Plex Movies
MOVIE_WISHLIST_SHEET=DVD Wish List
```

---

## ▶️ How to Use

1. Clone this repository
2. Create your `.env` file
3. Share both Plex Google Sheets (`Plex Movies`, `DVD Wish List`) with your service account email
4. Run the exporter:

```bash
python plex_catalog_exporter.py
```

5. To launch the web UI for editing the DVD Wish List:

```bash
python -m app.app
```

Open your browser to `http://localhost:5000` to view and edit the wishlist.

---

## 🔄 Sync Behavior

- Overwrites **each tab** in the Google Sheet matching the Excel sheets
- Extra tabs (e.g., `Notes`) in your Google Sheet are left untouched
- The **wishlist** is pulled live from Google Sheets at runtime

---

## 🧠 Backup Tags Logic

Backup types (`DVD`, `ISO`, `Blue-ray`, `Ripped`, `Backup`) are pulled from the **Labels** field in Plex metadata.

Add these labels to your Plex movies or episodes to track backup types. Multiple labels are supported per item.

---

## 📋 Roadmap

- [x] Replace local Wishlist tab with live data from Google Sheets
- [x] Automatically sync final Excel output to Google Sheets
- [x] Show bar chart of movie backup types
- [x] Add pie chart of TV episode backup coverage
- [x] Add TV Dashboard tab
- [x] Switch from Collections to Labels for backup tagging
- [x] Add web UI for viewing/editing the wish list
- [ ] Allow front-end to delete and re-order wish list items
- [ ] Format the Wishlist tab (freeze headers, auto-width)
- [ ] Auto-cleanup old timestamped folders after successful upload

---

## 📜 License

MIT License

You are free to use, modify, and distribute this tool with attribution.

---

## Author

**Erick Perales** — IT Architect, Cloud Migration Specialist  
[https://github.com/peralese](https://github.com/peralese)
