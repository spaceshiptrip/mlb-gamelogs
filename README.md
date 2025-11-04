# ESPN Play-by-Play Scraper

This tool converts ESPN MLB Play-by-Play data into a structured Excel file with detailed **At-Bat**, **Pitch**, and **Summary** sheets.

It supports both **static HTML files** and **live ESPN URLs**, automatically expands hidden pitch tables, detects **pitchers and pitching changes**, and records **scores** per play.

---

## 🧩 Features

- Parses ESPN Play-by-Play pages for MLB games
- Extracts **At-Bat**, **Pitch**, and **Summary** data
- Captures **pitch sequences** (pitch number, type, mph, result)
- Tracks **current pitcher** and detects **pitching changes**
- Records **away/home scores**
- Supports both **Playwright-rendered** and **saved HTML** sources
- Outputs a clean Excel workbook with 3 sheets

---

## ⚙️ Installation

### Option 1: Using `uv` (recommended)
```bash
uv tool install --with openpyxl pandas requests beautifulsoup4 playwright
playwright install chromium
```

### Option 2: Using `pip`
```bash
pip install -r requirements.txt
playwright install chromium
```

---

## 🧾 Full Requirements

Exact dependencies used in this project:

```
beautifulsoup4==4.14.2
certifi==2025.10.5
charset-normalizer==3.4.4
et_xmlfile==2.0.0
greenlet==3.2.4
idna==3.11
lxml==6.0.2
markdown-it-py==4.0.0
mdurl==0.1.2
numpy==2.3.4
openpyxl==3.1.5
pandas==2.3.3
playwright==1.55.0
pyee==13.0.0
Pygments==2.19.2
python-dateutil==2.9.0.post0
pytz==2025.2
requests==2.32.5
rich==14.2.0
six==1.17.0
soupsieve==2.8
typing_extensions==4.15.0
tzdata==2025.2
urllib3==2.5.0
xlsxwriter==3.2.9
```

---

## 🚀 Usage

### Parse a Live ESPN Game
```bash
python parse_atbats_to_excel.py --input "https://www.espn.com/mlb/playbyplay/_/gameId/401809302" --render -o game6.xlsx
```

### Parse a Saved HTML File
```bash
python parse_atbats_to_excel.py --input saved_playbyplay.html -o pbp.xlsx
```

**Tips:**
- If ESPN blocks requests (403), add `--render` to enable Chromium-based rendering.
- Use `--quiet` to reduce console output.

---

## 📊 Excel Output

**At-Bats sheet:**
| inning | half | batting_team | pitcher | description | away_score | home_score | pitch_count |
|--------|------|---------------|----------|--------------|-------------|-------------|--------------|

**Pitches sheet:**
| atbat_seq | inning | half | pitcher | pitch_no | result | pitch_type | mph |
|------------|--------|------|----------|-----------|----------|-------------|-----|

**Summary sheet:**
| seq | inning | half | pitcher | play | pitch_sequence |
|------|---------|------|----------|------|----------------|

---

## 🧠 Notes

- Playwright mode (`--render`) ensures full DOM visibility for hidden pitch tables.
- Pitcher names are inferred from ESPN’s text; pitching changes update automatically.
- If a page omits the pitcher, data still aligns with inning/half context.

---

## 🧰 Project Structure

```
espn-scraper/
├── parse_atbats_to_excel.py
├── requirements.txt
├── README.md
├── .gitignore
└── (generated) output.xlsx
```

---

**Author:** ChatGPT + Jay Torres  
**Version:** 1.2  
**Date:** 2025-11-03  
