#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import re
import sys
import time
from pathlib import Path
from typing import List, Tuple, Dict, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag

# ----------------------------
# Fetch / Render HTML
# ----------------------------

def fetch_html(src: str, quiet: bool = False) -> str:
    """
    If src looks like a URL, GET with headers; else treat as local file path.
    """
    if re.match(r'^https?://', src.strip(), re.I):
        if not quiet:
            print(f"Fetching URL: {src}")
        headers = {
            "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/120.0.0.0 Safari/537.36"),
            "Accept-Language": "en-US,en;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }
        r = requests.get(src, headers=headers, timeout=30)
        r.raise_for_status()
        return r.text
    else:
        p = Path(src)
        if not p.exists():
            raise FileNotFoundError(src)
        if not quiet:
            print(f"Reading file: {src}")
        return p.read_text(encoding="utf-8")

def render_with_playwright(url: str, quiet: bool = False, wait_ms: int = 1500) -> str:
    """
    Render dynamic ESPN page to materialize collapsed pitch tables.
    """
    if not quiet:
        print(f"Rendering URL (Playwright): {url}")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(user_agent=(
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ))
        page = ctx.new_page()
        page.goto(url, wait_until="domcontentloaded")
        # Expand all accordions if present
        page.wait_for_timeout(wait_ms)
        # Click all accordion headers to ensure inner tables exist in DOM
        headers = page.locator(".AtBatAccordion__header")
        count = headers.count()
        for i in range(count):
            try:
                btn = headers.nth(i)
                aria = btn.get_attribute("aria-expanded") or ""
                if aria.lower() != "true":
                    btn.click()
                    page.wait_for_timeout(50)
            except Exception:
                pass
        page.wait_for_timeout(500)
        html = page.content()
        browser.close()
        return html

# ----------------------------
# Parsing helpers
# ----------------------------

def text_or_none(node: Optional[Tag]) -> str:
    return node.get_text(strip=True) if node else ""

def find_all_atbats(soup: BeautifulSoup) -> List[Tag]:
    # ESPN uses <div class="AtBatAccordion"> per play
    return list(soup.select("div.AtBatAccordion"))

def parse_play_header(atbat_root: Tag) -> Tuple[str, Optional[int], Optional[int]]:
    """
    Return (description, away_score, home_score) from PlayHeader inside each AtBatAccordion.
    """
    desc = ""
    away = None
    home = None
    ph = atbat_root.select_one(".PlayHeader")
    if ph:
        d = ph.select_one(".PlayHeader__description")
        desc = text_or_none(d)
        a = ph.select_one(".PlayHeader__score--away")
        h = ph.select_one(".PlayHeader__score--home")
        try:
            away = int(text_or_none(a)) if a else None
        except ValueError:
            away = None
        try:
            home = int(text_or_none(h)) if h else None
        except ValueError:
            home = None
    return desc, away, home

def find_inning_context(cursor: Tag) -> Tuple[str, str, str]:
    """
    Walk up then left to find inning & half & batting team in nearby headers.
    Format returned inning as string number if possible; half in {'Top','Bottom','Mid','End'}.
    Team may be 'Unknown Team' if not found.
    """
    # Defaults
    inning = "Unknown"
    half = "Unknown"
    team = "Unknown Team"

    # Heuristic: look left (previous siblings) for containers with inning labels
    ptr = cursor
    for _ in range(3):  # climb a few levels
        if not ptr:
            break
        sib = ptr.previous_sibling
        hops = 0
        while sib and hops < 30:
            if isinstance(sib, Tag):
                txt = sib.get_text(" ", strip=True)
                if txt:
                    # Examples: "Top 3rd", "Bottom 1st", "Mid 7th", etc.
                    m = re.search(r'(?i)\b(Top|Bottom|Mid|End)\b\s+(\d+)(?:st|nd|rd|th)?', txt)
                    if m:
                        half = m.group(1).title()
                        inning = m.group(2)
                        # Try to find a team near this label
                        tm = re.search(r'(?i)\bvs\b\s+([A-Za-z .-]+)$', txt)
                        if tm:
                            team = tm.group(1).strip()
                        break
                    # Another pattern some ESPN pages show inside headers:
                    m2 = re.search(r'(?i)\b(Top|Bottom)\b.*?\bInning\b.*?(\d+)', txt)
                    if m2:
                        half = m2.group(1).title()
                        inning = m2.group(2)
                        break
            sib = sib.previous_sibling
            hops += 1
        if inning != "Unknown":
            break
        ptr = ptr.parent if isinstance(ptr, Tag) else None

    return inning, half, team

PITCH_RESULT_MAP = {
    "Strike Looking": "Strike Looking",
    "Strike Swinging": "Strike Swinging",
    "Foul Ball": "Foul Ball",
    "Ball": "Ball",
    "Hit By Pitch": "Hit By Pitch",
    "Single": "Single",
    "Double": "Double",
    "Triple": "Triple",
    "Home Run": "Home Run",
    "Pop Out": "Pop Out",
    "Fly Out": "Fly Out",
    "Ground Out": "Ground Out",
    "Batter Reached On Error - Batter To First": "Reached On Error",
}

def parse_pitch_table(body: Tag) -> List[Tuple[Optional[int], str, Optional[str], Optional[str]]]:
    """
    Each row is (pitch_number, result_text, pitch_type, mph)
    """
    rows = []
    table = body.select_one("table.Table")
    if not table:
        return rows
    for r in table.select("tbody.Table__TBODY > tr"):
        # first cell holds count icon & label
        pitch_no = None
        result = ""
        tds = r.select("td.Table__TD")
        if not tds:
            continue
        # Pitch number is in the little icon; result label is a sibling span
        count_icon = tds[0].select_one(".PitchCountIcon")
        if count_icon:
            try:
                pitch_no = int(count_icon.get_text(strip=True))
            except ValueError:
                pitch_no = None
        label_span = tds[0].select_one("span")
        result = text_or_none(label_span)

        pitch_type = text_or_none(tds[1]) if len(tds) > 1 else None
        mph = text_or_none(tds[2]) if len(tds) > 2 else None
        mph = mph if mph else None
        rows.append((pitch_no, result, pitch_type if pitch_type else None, mph))
    return rows

# ---- Pitcher tracking ----

PITCH_CHANGE_PATTERNS = [
    # "Pitching Change: Yency Almonte replaces Jose Berríos."
    re.compile(r'(?i)\bPitching Change:\s*([A-Z][a-zA-Z.\'-]+ [A-Z][a-zA-Z.\'-]+)\s+replaces\s+([A-Z][a-zA-Z.\'-]+ [A-Z][a-zA-Z.\'-]+)'),
    # "X relieved by Y"
    re.compile(r'(?i)\b([A-Z][a-zA-Z.\'-]+ [A-Z][a-zA-Z.\'-]+)\s+relieved by\s+([A-Z][a-zA-Z.\'-]+ [A-Z][a-zA-Z.\'-]+)'),
    # "X replaces Y"
    re.compile(r'(?i)\b([A-Z][a-zA-Z.\'-]+ [A-Z][a-zA-Z.\'-]+)\s+replaces\s+([A-Z][a-zA-Z.\'-]+ [A-Z][a-zA-Z.\'-]+)'),
]

def detect_pitching_change(desc: str) -> Optional[str]:
    """
    Return the NEW pitcher name if a pitching change is described; else None.
    """
    for pat in PITCH_CHANGE_PATTERNS:
        m = pat.search(desc)
        if m:
            # Prefer 'new' as first group if pattern is "Pitching Change: NEW replaces OLD"
            # Otherwise infer new as the second group.
            if 'Pitching Change' in pat.pattern:
                return f"{m.group(1)}"
            # for 'relieved by' or 'replaces'
            return f"{m.group(2)}"
    return None

def sniff_pitcher_from_context(atbat_root: Tag) -> Optional[str]:
    """
    Some pages show 'Pitching: Name' nearby. Search the at-bat cluster.
    """
    # 1) in PlayHeader or siblings?
    txt = atbat_root.get_text(" ", strip=True)
    m = re.search(r'(?i)\bPitching:\s*([A-Z][a-zA-Z.\'-]+ [A-Z][a-zA-Z.\'-]+)', txt)
    if m:
        return m.group(1).strip()
    # 2) less reliable global match left here as fallback
    return None

# ----------------------------
# Main parse
# ----------------------------

def parse_play_by_play(html: str, verbose: bool = True) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    soup = BeautifulSoup(html, "html.parser")
    atbats = find_all_atbats(soup)
    if verbose:
        print(f"Found {len(atbats)} at-bats on the page.")

    atbat_rows = []
    pitch_rows = []
    summary_rows = []

    # Track pitcher per half-inning (since changes reset with half)
    current_pitcher_by_half: Dict[Tuple[str, str], Optional[str]] = {}

    for idx, ab in enumerate(atbats, start=1):
        desc, away_sc, home_sc = parse_play_header(ab)
        body = ab.select_one(".AtBatAccordion__body")

        inning, half, batting_team = find_inning_context(ab)
        half_key = (inning, half)

        # update pitcher if discoverable in this block
        sniffed = sniff_pitcher_from_context(ab)
        if sniffed:
            current_pitcher_by_half[half_key] = sniffed

        # If the description itself signals a change, update BEFORE recording this at-bat
        new_pitcher = detect_pitching_change(desc)
        if new_pitcher:
            current_pitcher_by_half[half_key] = new_pitcher

        pitcher = current_pitcher_by_half.get(half_key)

        # Extract inner pitch table
        inner = []
        if body:
            inner = parse_pitch_table(body)

        # Status output
        if verbose:
            inn_show = inning if inning else "Unknown Inning"
            tm_show = batting_team if batting_team else "Unknown Team"
            print(f"Parsing Inning {inn_show} ({half} – {tm_show}): {desc}")
            if inner:
                joined = ", ".join([f"P{p[0] if p[0] is not None else '?'}: {p[1]}" for p in inner])
                print(f"   • Captured {len(inner)} pitch rows → {joined}")
            else:
                print("   • Captured 0 pitch rows")

        # Record at-bat row
        atbat_rows.append({
            "seq": idx,
            "inning": inning,
            "half": half,
            "batting_team": batting_team,
            "description": desc,
            "pitcher": pitcher if pitcher else "",
            "away_score": away_sc,
            "home_score": home_sc,
            "pitch_count": len(inner)
        })

        # Record pitch rows (one per pitch)
        for pno, pres, ptype, mph in inner:
            pitch_rows.append({
                "atbat_seq": idx,
                "inning": inning,
                "half": half,
                "batting_team": batting_team,
                "pitcher": pitcher if pitcher else "",
                "pitch_no": pno,
                "result": pres,
                "pitch_type": ptype,
                "mph": mph,
                "away_score": away_sc,
                "home_score": home_sc,
                "atbat_desc": desc
            })

        # Keep a slim summary (for convenience)
        if inner:
            pitch_seq = " • ".join([f"P{p[0] if p[0] else '?'}: {p[1]}" for p in inner])
        else:
            pitch_seq = ""
        summary_rows.append({
            "seq": idx,
            "inning": inning,
            "half": half,
            "batting_team": batting_team,
            "pitcher": pitcher if pitcher else "",
            "play": desc,
            "pitch_sequence": pitch_seq,
            "away_score": away_sc,
            "home_score": home_sc
        })

        # If this AB is a pitching change with no batter action, continue; otherwise,
        # also handle cases where the next AB should still carry the updated pitcher.
        if new_pitcher and verbose:
            print(f"   • Pitching change detected → current pitcher now: {new_pitcher}")

    atbats_df = pd.DataFrame(atbat_rows)
    pitches_df = pd.DataFrame(pitch_rows)
    summary_df = pd.DataFrame(summary_rows)
    return atbats_df, pitches_df, summary_df

# ----------------------------
# CLI
# ----------------------------

def main():
    ap = argparse.ArgumentParser(description="Parse ESPN MLB Play-by-Play into Excel with At-Bats & Pitches.")
    ap.add_argument("--input", "-i", required=True, help="URL to ESPN play-by-play or path to saved HTML.")
    ap.add_argument("--output", "-o", default="espn_playbyplay.xlsx", help="Output Excel filename.")
    ap.add_argument("--render", action="store_true", help="Use Playwright to render the page (dynamic).")
    ap.add_argument("--quiet", action="store_true", help="Reduce console output.")
    args = ap.parse_args()

    src = args.input
    if re.match(r'^https?://', src.strip(), re.I) and args.render:
        html = render_with_playwright(src, quiet=args.quiet)
    else:
        html = fetch_html(src, quiet=args.quiet)

    atbats_df, pitches_df, summary_df = parse_play_by_play(html, verbose=(not args.quiet))

    out_path = Path(args.output)
    if not args.quiet:
        print(f"Writing Excel → {out_path.resolve()}")

    # Ensure openpyxl exists (friendly msg if not)
    try:
        import openpyxl  # noqa: F401
    except Exception:
        print("openpyxl is required for Excel output. Install with:\n  uv tool install --with openpyxl pandas\nor\n  pip install openpyxl", file=sys.stderr)
        sys.exit(1)

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        # Preserve useful column order
        atbat_cols = ["seq","inning","half","batting_team","pitcher","description","away_score","home_score","pitch_count"]
        pitch_cols  = ["atbat_seq","inning","half","batting_team","pitcher","pitch_no","result","pitch_type","mph","away_score","home_score","atbat_desc"]
        summ_cols   = ["seq","inning","half","batting_team","pitcher","play","pitch_sequence","away_score","home_score"]

        if not atbats_df.empty:
            atbats_df = atbats_df.reindex(columns=atbat_cols)
        if not pitches_df.empty:
            pitches_df = pitches_df.reindex(columns=pitch_cols)
        if not summary_df.empty:
            summary_df = summary_df.reindex(columns=summ_cols)

        atbats_df.to_excel(xw, sheet_name="At-Bats", index=False)
        pitches_df.to_excel(xw, sheet_name="Pitches", index=False)
        summary_df.to_excel(xw, sheet_name="Summary", index=False)

    if not args.quiet:
        print("Done.")

if __name__ == "__main__":
    main()

