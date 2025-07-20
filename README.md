# Plex Catalog Exporter

This Python utility connects to your Plex server and exports your movie and TV show libraries into an organized Excel workbook. It includes a dashboard summary, individual sheets per movie library, a combined TV shows sheet, and a wishlist tab for tracking future acquisitions.

---

## 📦 Features

- 📁 One worksheet per **movie** library section (e.g., Movies, Classics)
- 🎬 One combined **TV_Shows** worksheet with detailed episode breakdowns
- 📊 A **Dashboard** tab summarizing movie stats and backup coverage
- ✅ **Backup detection** using **Plex Collections**:
  - Collections like `Backup`, `ISO`, `DVD`, `Blue-ray` are used to determine backup status
- 🔄 **Fallback logic**: if no collections are found, it checks the file path for `.iso`, `.vob`, or `dvd` folder names
- 📂 Output saved in a **timestamped folder** under `output/`
- 📈 Totals row with **% Backed Up** on every sheet
- 📝 A **Wishlist** sheet for tracking missing content

---

## 📁 Example Excel Layout

### 🎬 Movie Sheet (`Movies`, `Classics`, etc.)

| Title          | Backup | Type     | Path                          |
|----------------|--------|----------|-------------------------------|
| Casablanca     | Yes    | ISO      | /plex/movies/casablanca.iso   |
| Ronin          | Yes    | ISO, DVD | /plex/movies/ronin.iso        |
| Inception      | No     |          | /plex/movies/inception.mkv    |
| **Total**      | 2      |          |                               |

### 📺 TV_Shows Sheet

| Show Title     | Season | Episode | Episode Title | Backup | Type | Path                         |
|----------------|--------|---------|----------------|--------|------|------------------------------|
| The Office     | 1      | 1       | Pilot          | Yes    | ISO  | /plex/tv/office/s01e01.iso   |
| Breaking Bad   | 2      | 3       | Bit by a Bee   | No     |      | /plex/tv/breakingbad/s02e03.mkv |
| **Total**      |        |         |                | 1      |      |                              |

### 📊 Dashboard

| Category   | Total Movies | With Backup | % Backed Up |
|------------|---------------|--------------|--------------|
| Movies     | 200           | 90           | 45.0         |
| Classics   | 50            | 30           | 60.0         |
| **Total**  | 250           | 120          | 48.0         |

---

## 🚀 Setup & Usage

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/plex-catalog-exporter.git
cd plex-catalog-exporter
```

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
- Optional: Use **Collections** in Plex (e.g., Backup, ISO, DVD) for richer reporting

---

## 🧩 Future Enhancements

- TV Dashboard summary
- Allow Multiple collections to be combined (e.g., `ISO, DVD`)
- Include IMDb/TMDb ratings
- Column auto-sizing and formatting
- GUI or web-based launcher
- Conditional formatting for backup status

---

## 📄 License

MIT License © 2025 Erick Perales  
You are free to use, modify, and distribute this tool with attribution.

---

## Author

**Erick Perales** — IT Architect, Cloud Migration Specialist
