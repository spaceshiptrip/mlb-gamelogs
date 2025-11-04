# ESPN Play-by-Play Scraper

This script parses **ESPN MLB Play-by-Play** pages — either directly from a URL or from a saved HTML file — 
and extracts detailed **At-Bat**, **Pitch**, and **Summary** data into an Excel workbook.

It supports ESPN’s **dynamic pitch tables** (those hidden behind collapsible "At Bat" accordions) by using 
**Playwright** for JavaScript rendering.

---

## 🧠 Features

- Extracts every **At-Bat** and each individual **Pitch** (pitch number, result, type, mph)
- Tracks **Pitchers**, including automatic updates for **pitching changes**
- Captures **scores** (away/home) at each play
- Works with both **static HTML files** and **live ESPN URLs**
- Saves results into an Excel file with three sheets:
  - `At-Bats`
  - `Pitches`
  - `Summary`

---

## ⚙️ Installation

You’ll need Python 3.9+ and a few packages:

```bash
uv tool install --with openpyxl pandas requests beautifulsoup4 playwright
playwright install chromium
```

Or using `pip`:

```bash
pip install pandas requests beautifulsoup4 openpyxl playwright
playwright install chromium
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

If ESPN blocks requests (403 Forbidden), use `--render` to enable Playwright, which runs Chromium to load the full page.

---

## 📄 Output Example

**At-Bats sheet:**
| inning | half  | batting_team | pitcher | description | away_score | home_score | pitch_count |
|--------|-------|---------------|----------|--------------|-------------|-------------|--------------|
| 1 | Top | Dodgers | Berríos | Ohtani struck out swinging | 0 | 0 | 6 |

**Pitches sheet:**
| atbat_seq | inning | half | pitcher | pitch_no | result | pitch_type | mph |
|------------|--------|------|----------|-----------|----------|-------------|-----|
| 1 | 1 | Top | Berríos | 1 | Strike Swinging | Four-seam FB | 96 |

**Summary sheet:**
| seq | inning | half | pitcher | play | pitch_sequence |
|------|---------|------|----------|------|----------------|
| 1 | 1 | Top | Berríos | Ohtani struck out swinging | P1: Strike Swinging, P2: Strike Swinging, P3: Ball, P4: Ball, P5: Foul Ball, P6: Strike Swinging |

---

## 🧩 Notes

- Playwright mode (`--render`) is slower but ensures all hidden pitch tables load.
- Scores and pitcher names depend on how ESPN structures that specific game’s page.
- If a pitcher isn’t detected automatically, it will still capture all pitches under the correct inning/half.

---

**Author:** ChatGPT + Jay Torres  
**Date:** 2025-11-03  
**Version:** 1.0
