# Plex Catalog Exporter

This Python script exports your Plex movie and TV libraries into an organized Excel spreadsheet, including dashboards, backup tracking, and a wish list integration from Google Sheets.

---

## 📦 Features

- 🔍 Exports **one Excel tab per movie library** (e.g., Movies, Classics, Anime)
- 📺 Generates a **TV Shows summary** sheet and a **TV Dashboard** with charts
- 📊 Creates a **Dashboard** tab summarizing movie backup stats by type
- ☁️ **Automatically uploads the Excel workbook to Google Sheets**
- 🌐 **Pulls a movie wishlist** from an external Google Sheet
- 📁 Saves exports in timestamped folders (e.g., `output/catalog_2025-07-21_130022/`)

---

## 📂 Output Structure

Each Excel export includes:

| Sheet Name       | Description                                  |
|------------------|----------------------------------------------|
| `Dashboard`      | Backup summary per movie library             |
| `TV_Dashboard`   | TV shows backup coverage + chart             |
| `Movies`, `Classics`, etc. | One tab per Plex movie library       |
| `TV_Shows`       | Combined list of all TV shows                |
| `Wishlist`       | Pulled live from external Google Sheet       |

---

## ✅ Requirements

- Python 3.9+
- A Plex Media Server
- A service account with access to the desired Google Sheets
- `.env` file with the following:

```
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
4. Run:

```bash
python plex_catalog_exporter.py
```

---

## 🔄 Sync Behavior

- Overwrites **each tab** in the Google Sheet matching the Excel sheets
- Extra tabs (e.g., `Notes`) in your Google Sheet are left untouched

---

## 🧠 Backup Tags Logic

Backup types (e.g., DVD, ISO, Ripped, Blue-ray) are pulled from **Labels**, not Collections, in Plex.

Supported tags:
- `DVD`, `ISO`, `Blue-ray`, `Ripped`, `Backup`

These labels can be added to each movie or episode in Plex, and the script will detect them to classify backups.

---

## 📋 Roadmap

- [x] Replace local Wishlist tab with live data from Google Sheets
- [x] Automatically sync final Excel output to Google Sheets
- [x] Show bar chart of movie backup types
- [x] Add pie chart of TV episode backup coverage
- [x] Add TV Dashboard tab
- [ ] **Add front-end to edit the movie wishlist directly**
- [ ] Add automatic cleanup step to remove timestamped folders (optional)
- [ ] Add formatting for the wishlist sheet (freeze headers, auto-width)
- [ ] Validate that movie types in wishlist are `Movie` before inserting

---

## 📜 License

MIT License

You are free to use, modify, and distribute this tool with attribution.

---

## Author

**Erick Perales** — IT Architect, Cloud Migration Specialist
[https://github.com/peralese](https://github.com/peralese)