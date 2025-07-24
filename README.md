# Plex Catalog Exporter

This Python utility connects to your Plex server and exports your movie and TV show libraries into an organized Excel workbook. It includes a dashboard summary, individual sheets per movie library, a combined TV shows sheet, a TV dashboard, and a wishlist tab for tracking future acquisitions.

---

## 📦 Features

- 📁 One worksheet per **movie** library section (e.g., Movies, Classics)
- 🎬 One combined **TV_Shows** worksheet with detailed episode breakdowns
- 📊 A **Dashboard** tab summarizing movie stats and backup coverage
- 📺 A **TV_Dashboard** tab summarizing per-show episode backup stats
- ✅ **Backup detection** using **Plex Labels** (not Collections):
  - Labels like `Backup`, `ISO`, `DVD`, `Blue-ray`, `Ripped` are used to determine backup status
- 🔄 **Fallback logic**: if no labels are found, it checks the file path for `.iso`, `.vob`, or `dvd` folder names
- 📈 Totals row with **% Backed Up** on every sheet
- 📊 Dashboard includes a **bar chart** showing total counts per Backup Type (DVD, Blue-ray, ISO, Ripped)
- 📈 TV Dashboard includes a **pie chart** showing % of TV episodes backed up vs not backed up
- 📂 Output saved in a **timestamped folder** under `output/`
- 📝 A **Wishlist** sheet for tracking missing content

---

## 📁 Example Excel Layout

### 🎬 Movie Sheet (`Movies`, `Classics`, etc.)

| Title        | Backup | Type     | Path                        |
|--------------|--------|----------|-----------------------------|
| Casablanca   | Yes    | ISO      | /plex/movies/casablanca.iso |
| Ronin        | Yes    | ISO, DVD | /plex/movies/ronin.iso      |
| Inception    | No     |          | /plex/movies/inception.mkv  |
| **Total**    | 2      |          |                             |

### 📺 TV_Shows Sheet

| Show Title     | Season | Episode | Episode Title | Backup | Type | Path                         |
|----------------|--------|---------|----------------|--------|------|------------------------------|
| The Office     | 1      | 1       | Pilot          | Yes    | ISO  | /plex/tv/office/s01e01.iso   |
| Breaking Bad   | 2      | 3       | Bit by a Bee   | No     |      | /plex/tv/breakingbad/s02e03.mkv |
| **Total**      |        |         |                | 1      |      |                              |

### 📊 Dashboard

| Category | Movie Count | Backup Type (DVD, Blue-ray, ISO, Ripped) | % Backed Up |
|----------|--------------|-------------------------------------------|--------------|
| Movies   | 200          | 50 / 10 / 20 / 5                         | 42.5         |
| Classics | 50           | 5 / 2 / 3 / 1                            | 22.0         |
| **Total**| 250          | 55 / 12 / 23 / 6                         | 38.4         |

### 📺 TV_Dashboard

| Show Title     | Total Episodes | Backed Up | % Backed Up |
|----------------|----------------|-----------|--------------|
| The Office     | 200            | 150       | 75.0         |
| Breaking Bad   | 62             | 62        | 100.0        |
| Stranger Things| 34             | 20        | 58.8         |
| **Total**      | 296            | 232       | 78.4         |

---

## 🚀 Setup & Usage

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/plex-catalog-exporter.git
cd plex-catalog-exporter

### 2. Install Dependencies

It’s recommended to use a virtual environment:

```bash
pip install -r requirements.txt
```

> Required packages: `plexapi`, `pandas`, `openpyxl`, `python-dotenv`

### 3. Create a `.env` File

Create a `.env` file in the root of the project with the following:

```
PLEX_BASEURL=http://localhost:32400
PLEX_TOKEN=your_plex_token_here
IGNORE_LIBRARIES=Music Videos
```

> You can exclude specific Plex libraries by listing them under `IGNORE_LIBRARIES`, separated by commas.

To get your Plex token:  
[📖 Plex Token Guide](https://support.plex.tv/articles/204059436-finding-an-authentication-token-x-plex-token/)

### 4. Run the Script

```bash
python plex_catalog_exporter.py
```

The output Excel file will be created in:

```
output/YYYY-MM-DD_HH-MM-SS/plex_media_catalog.xlsx
```

---

## 🛠️ Requirements

- Python 3.7+
- Access to your Plex server and a valid Plex Token
- Required: Use **Labels** in Plex (e.g., Backup, ISO, DVD, Blue-ray, Ripped) instead of Collections for tagging

---

## 🧩 Future Enhancements

- Support multiple backup types on one item (e.g., `ISO, DVD`)
- Include IMDb/TMDb ratings
- GUI or web-based launcher
- Conditional formatting
- Push output to Google Sheets
- Enhanced charting (e.g. color-coded bar chart)

---

## 📄 License

MIT License © 2025 Erick Perales  
You are free to use, modify, and distribute this tool with attribution.

---

## Author

**Erick Perales** — IT Architect, Cloud Migration Specialist
[https://github.com/peralese](https://github.com/peralese)