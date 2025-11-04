#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
ESPN MLB Play-by-Play → Excel (ALL pitch tables)
- If --render is used with an ESPN URL, launches Playwright (headless Chromium),
  auto-expands EVERY AtBatAccordion, waits for pitch tables to load, then parses.
- If --render is omitted, parses the given URL/file as-is (works only if all pitch tables are present).
- Status logs like "Parsing Inning 3 (Top – TOR): ..." and "• Captured N pitch rows → P1: Ball, ..."
- Outputs 3 sheets: AtBats, Pitches, GameSummary.

Deps: playwright, beautifulsoup4, lxml, requests, pandas, openpyxl
"""

import argparse
import os
import re
import sys
import time
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag

# Playwright is optional (only needed with --render)
try:
    from playwright.sync_api import sync_playwright
except Exception:
    sync_playwright = None

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Cache-Control": "no-cache",
}

ORDINAL_MAP = {
    "1st": 1, "2nd": 2, "3rd": 3, "4th": 4, "5th": 5, "6th": 6,
    "7th": 7, "8th": 8, "9th": 9, "10th": 10, "11th": 11, "12th": 12,
    "13th": 13, "14th": 14, "15th": 15, "16th": 16, "17th": 17, "18th": 18,
}

def is_url(s: str) -> bool:
    return s.startswith("http://") or s.startswith("https://")

def fetch_html_via_requests(url: str, quiet: bool = False) -> str:
    if not quiet:
        print(f"Fetching URL (requests): {url}")
    sess = requests.Session()
    for attempt in range(4):
        try:
            resp = sess.get(url, headers=DEFAULT_HEADERS, timeout=25)
            if resp.status_code == 403:
                # try again with small tweaks
                time.sleep(0.8 + attempt * 0.4)
                resp = sess.get(
                    url,
                    headers={**DEFAULT_HEADERS, "Pragma": "no-cache", "DNT": "1"},
                    timeout=25,
                )
            resp.raise_for_status()
            return resp.text
        except Exception as e:
            if attempt == 3:
                raise
            if not quiet:
                print(f"  retry {attempt+1}/4 after: {e}")
            time.sleep(1.2)

def fetch_html(src: str, quiet: bool = False, save_html: Optional[str] = None, render: bool = False) -> str:
    """
    - If render=True and src is a URL → use Playwright to expand all at-bats and return full HTML.
    - Else if src is a URL → GET via requests.
    - Else read local file.
    """
    if is_url(src) and render:
        if sync_playwright is None:
            raise RuntimeError("Playwright not installed. Run: python -m playwright install")
        if not quiet:
            print(f"Rendering URL (Playwright): {src}")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(user_agent=DEFAULT_HEADERS["User-Agent"])
            page = ctx.new_page()
            page.set_default_timeout(30000)
            page.goto(src, wait_until="domcontentloaded")

            # Some pbp pages lazy-load sections while scrolling – scroll to the bottom slowly
            # to force-load all accordions.
            last_height = 0
            for _ in range(8):
                page.evaluate("window.scrollBy(0, document.body.scrollHeight)")
                time.sleep(0.6)
                h = page.evaluate("document.body.scrollHeight")
                if h == last_height:
                    break
                last_height = h

            # Expand ALL at-bat accordions (they are buttons with class .AtBatAccordion__header)
            # We attempt multiple passes to catch chunks that load late.
            for _ in range(3):
                buttons = page.query_selector_all("button.AtBatAccordion__header, .AtBatAccordion > button")
                for btn in buttons:
                    try:
                        expanded = btn.get_attribute("aria-expanded")
                        if expanded != "true":
                            btn.click()
                            # Wait briefly for the body to appear
                            time.sleep(0.05)
                    except Exception:
                        pass
                time.sleep(0.3)

            # After expanding, wait until most pitch tables are present (best-effort heuristic)
            # This prevents racing on slower connections.
            for _ in range(10):
                count = page.evaluate("document.querySelectorAll('div.Collapse.AtBatAccordion__body .PitchTable').length")
                # break early if we have a decent number
                if count and count > 20:
                    break
                time.sleep(0.3)

            html = page.content()
            browser.close()

        if save_html:
            with open(save_html, "w", encoding="utf-8") as f:
                f.write(html)
        return html

    if is_url(src):
        html = fetch_html_via_requests(src, quiet=quiet)
        if save_html:
            with open(save_html, "w", encoding="utf-8") as f:
                f.write(html)
        return html

    # Local file
    if not quiet:
        print(f"Reading file: {src}")
    with open(src, "r", encoding="utf-8") as f:
        return f.read()

def inner_text(el: Optional[Tag]) -> str:
    return el.get_text(" ", strip=True) if el else ""

def _safe_int(x):
    try:
        return int(str(x).strip())
    except Exception:
        return None

# ---------- Team header (best-effort pretty logs) ----------
def parse_game_header_teams(soup: BeautifulSoup) -> Tuple[Optional[str], Optional[str]]:
    away = home = None
    og_title = soup.select_one('meta[property="og:title"]')
    if og_title and og_title.get("content"):
        m = re.search(r'^(.*?) vs\. (.*?) -', og_title["content"])
        if m:
            away, home = m.group(1).strip(), m.group(2).strip()
            return away, home
    sb = soup.select_one('[data-analytics="scoreboard"], .Scoreboard')
    if sb:
        abbrs = [el.get_text(strip=True) for el in sb.select(".ScoreCell__TeamName, .ScoreCell__Abbrev")]
        if len(abbrs) >= 2:
            away = away or abbrs[0]
            home = home or abbrs[1]
    return away, home

# ---------- Inning context parsing ----------
def nearest_inning_header_text(node: Tag) -> str:
    # Check previous siblings
    sib = node.previous_sibling
    while isinstance(sib, Tag):
        txt = inner_text(sib)
        if re.search(r'\b(Top|Bottom)\b', txt, re.I) and re.search(r'\b\d{1,2}(?:st|nd|rd|th)\b', txt, re.I):
            return txt
        if sib.select_one(".InningHeader, .Accordion__header"):
            t2 = inner_text(sib)
            if t2:
                return t2
        sib = sib.previous_sibling

    # Check ancestors direct child headers
    parent = node.parent
    while isinstance(parent, Tag):
        headers = parent.find_all(["h1", "h2", "h3", "header"], recursive=False)
        for h in headers:
            txt = inner_text(h)
            if re.search(r'\b(Top|Bottom)\b', txt, re.I) and re.search(r'\b\d{1,2}(?:st|nd|rd|th)\b', txt, re.I):
                return txt
        hnode = parent.select_one(":scope > .InningHeader, :scope > .Accordion__header")
        if hnode:
            return inner_text(hnode)
        parent = parent.parent
    return ""

def parse_inning_half_team(header_text: str, away_team: Optional[str], home_team: Optional[str]) -> Tuple[Optional[int], Optional[str], Optional[str]]:
    half = None
    inn = None
    team = None

    m_half = re.search(r'\b(Top|Bottom)\b', header_text, re.I)
    if m_half:
        half = m_half.group(1).title()
    m_ord = re.search(r'\b(\d{1,2}(?:st|nd|rd|th))\b', header_text)
    if m_ord:
        inn = ORDINAL_MAP.get(m_ord.group(1))
    m_team = re.search(r'[–-]\s*([A-Za-z.\s-]{2,})$', header_text)  # text after dash/en-dash
    if m_team:
        team = m_team.group(1).strip()
    if not team and half:
        team = (away_team if half.lower() == "top" else home_team) or None
    return inn, half, team

# ---------- Core parsing ----------
def build_atbat_maps(soup: BeautifulSoup) -> Tuple[Dict[str, dict], Dict[str, str]]:
    """
    Build:
      atbat_info[atbat_key] = { 'desc','away_score','home_score','button','li','body_id' }
      body_to_atbat[body_id] = atbat_key
    """
    atbat_info: Dict[str, dict] = {}
    body_to_atbat: Dict[str, str] = {}

    # All headers (buttons)
    buttons = soup.select("button.AtBatAccordion__header, .AtBatAccordion > button")
    for btn in buttons:
        atbat_key = btn.get("id")
        if not atbat_key:
            continue
        body_id = btn.get("aria-controls")  # often "<id>-pitches"

        desc = inner_text(btn.select_one(".PlayHeader__description")) or inner_text(btn)
        away_score = inner_text(btn.select_one(".PlayHeader__score--away"))
        home_score = inner_text(btn.select_one(".PlayHeader__score--home"))

        info = {
            "desc": desc,
            "away_score": away_score,
            "home_score": home_score,
            "button": btn,
            "li": btn.find_parent("li"),
            "body_id": body_id,
        }
        atbat_info[atbat_key] = info
        if body_id:
            body_to_atbat[body_id] = atbat_key

    # Bodies → which at-bat?
    bodies = soup.select("div.Collapse.AtBatAccordion__body")
    for body in bodies:
        b_id = body.get("id")
        labeled_by = body.get("aria-labelledby")
        if labeled_by and labeled_by in atbat_info:
            if b_id:
                body_to_atbat[b_id] = labeled_by
            if not atbat_info[labeled_by].get("body_id") and b_id:
                atbat_info[labeled_by]["body_id"] = b_id

    return atbat_info, body_to_atbat

def parse_pitch_body(body: Tag) -> List[dict]:
    rows: List[dict] = []
    trs = body.select("tbody tr")
    for tr in trs:
        tds = tr.find_all("td")
        if len(tds) < 6:
            continue

        # Pitch # + label
        pitch_no = None
        result_txt = None
        icon = tds[0].select_one(".PitchCountIcon")
        if icon and icon.text:
            pitch_no = _safe_int(icon.text)
        spans = tds[0].find_all("span")
        if spans:
            result_txt = spans[-1].get_text(strip=True)

        pitch_type = inner_text(tds[1]) if len(tds) > 1 else None
        mph = _safe_int(inner_text(tds[2]) if len(tds) > 2 else None)

        # Zone
        zone = None
        if len(tds) > 3:
            hz = tds[3].select_one(".HitzoneIcon__location")
            if hz:
                for c in hz.get("class", []):
                    if c.startswith("HitzoneIcon__location--"):
                        zone = c.split("--", 1)[1]
                        break

        # Bases
        on1b = on2b = on3b = False
        if len(tds) > 4:
            on1b = bool(tds[4].select_one(".diamond.first-base.is--active"))
            on2b = bool(tds[4].select_one(".diamond.second-base.is--active"))
            on3b = bool(tds[4].select_one(".diamond.third-base.is--active"))

        # Field (style contains xy relative position)
        field_style = None
        if len(tds) > 5:
            fld = tds[5].select_one(".PlayFieldIcon__location")
            if fld:
                field_style = fld.get("style")

        rows.append({
            "pitch_no": pitch_no,
            "result": result_txt,
            "pitch_type": pitch_type,
            "mph": mph,
            "zone": zone,
            "on1b": on1b,
            "on2b": on2b,
            "on3b": on3b,
            "field_style": field_style,
        })

    rows.sort(key=lambda r: (9999 if r["pitch_no"] is None else r["pitch_no"]))
    return rows

def parse_all_pitches(soup: BeautifulSoup, body_to_atbat: Dict[str, str]) -> Tuple[Dict[str, List[dict]], Dict[str, str]]:
    pitch_map: Dict[str, List[dict]] = {}
    pitch_summary: Dict[str, str] = {}

    # Note: after Playwright expansion, every body should contain .PitchTable
    bodies = soup.select("div.Collapse.AtBatAccordion__body")
    for body in bodies:
        body_id = body.get("id")
        if not body_id:
            continue

        atbat_key = body_to_atbat.get(body_id)
        if not atbat_key and body_id.endswith("-pitches"):
            atbat_key = body_id[:-8]

        rows = parse_pitch_body(body)
        if not rows:
            continue

        key = atbat_key or body_id
        pitch_map.setdefault(key, [])
        for r in rows:
            rr = dict(r)
            rr["atbat_key"] = key
            pitch_map[key].append(rr)

    # Build summaries
    for k, rows in pitch_map.items():
        seq = []
        for r in rows:
            tag = f'P{r["pitch_no"]}' if r["pitch_no"] is not None else "P?"
            lab = r["result"] or ""
            seq.append(f"{tag}: {lab}")
        pitch_summary[k] = " • Captured {} pitch rows → ".format(len(rows)) + ", ".join(seq)
    return pitch_map, pitch_summary

def parse_play_by_play(full_html: str, verbose: bool=False) -> Tuple[List[dict], List[dict], dict]:
    soup = BeautifulSoup(full_html, "lxml")

    away_team, home_team = parse_game_header_teams(soup)
    atbat_info, body_to_atbat = build_atbat_maps(soup)
    pitch_map, pitch_summary = parse_all_pitches(soup, body_to_atbat)

    atbats: List[dict] = []
    for atbat_key, info in atbat_info.items():
        li = info.get("li")
        desc = info.get("desc", "")
        away_score = info.get("away_score", "")
        home_score = info.get("home_score", "")
        body_id = info.get("body_id")

        ctx_text = nearest_inning_header_text(li) if isinstance(li, Tag) else ""
        inning_num, half, team_name = parse_inning_half_team(ctx_text, away_team, home_team)

        if verbose:
            ih = inning_num if inning_num is not None else "Unknown"
            hh = half or "Unknown"
            tt = team_name or "Unknown Team"
            print(f"Parsing Inning {ih} ({hh} – {tt}): {desc}")
            seq = pitch_summary.get(atbat_key) or (pitch_summary.get(body_id) if body_id else "")
            if seq:
                print("  " + seq)
            else:
                print("  • Captured 0 pitch rows")

        atbats.append({
            "atbat_key": atbat_key,
            "inning": inning_num,
            "half": half,
            "team": team_name,
            "description": desc,
            "away_score": away_score,
            "home_score": home_score,
            "pitch_sequence": (pitch_summary.get(atbat_key) or (pitch_summary.get(body_id) if body_id else "")).replace("• ", ""),
        })

    pitch_rows: List[dict] = []
    for k, rows in pitch_map.items():
        for r in rows:
            pitch_rows.append({
                "atbat_key": r.get("atbat_key") or k,
                "pitch_no": r.get("pitch_no"),
                "result": r.get("result"),
                "pitch_type": r.get("pitch_type"),
                "mph": r.get("mph"),
                "zone": r.get("zone"),
                "on1b": r.get("on1b"),
                "on2b": r.get("on2b"),
                "on3b": r.get("on3b"),
                "field_style": r.get("field_style"),
            })

    summary = {
        "away_team": away_team,
        "home_team": home_team,
        "total_atbats_found": len(atbats),
        "total_pitch_events_found": len(pitch_rows),
    }

    return atbats, pitch_rows, summary

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="Scrape ESPN MLB play-by-play (all at-bat pitch details) to Excel.")
    ap.add_argument("--input", "-i", required=True, help="ESPN play-by-play URL or local HTML file")
    ap.add_argument("--output", "-o", default="pbp.xlsx", help="Output Excel filename (default: pbp.xlsx)")
    ap.add_argument("--quiet", "-q", action="store_true", help="Reduce console output")
    ap.add_argument("--save-html", help="Also save fetched/expanded HTML to this path (debug)")
    ap.add_argument("--render", action="store_true", help="Use Playwright to expand all at-bats (required for ESPN URLs to capture ALL pitch tables)")
    args = ap.parse_args()

    html = fetch_html(args.input, quiet=args.quiet, save_html=args.save_html, render=args.render)
    atbats, pitches, summary = parse_play_by_play(html, verbose=(not args.quiet))

    df_atbats = pd.DataFrame(atbats, columns=[
        "atbat_key", "inning", "half", "team",
        "description", "away_score", "home_score", "pitch_sequence"
    ])
    df_pitches = pd.DataFrame(pitches, columns=[
        "atbat_key", "pitch_no", "result", "pitch_type", "mph", "zone",
        "on1b", "on2b", "on3b", "field_style"
    ])
    df_summary = pd.DataFrame([summary])

    out_path = os.path.abspath(args.output)
    if not args.quiet:
        print(f"Writing Excel → {out_path}")

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        df_atbats.to_excel(xw, sheet_name="AtBats", index=False)
        df_pitches.to_excel(xw, sheet_name="Pitches", index=False)
        df_summary.to_excel(xw, sheet_name="GameSummary", index=False)

    if not args.quiet:
        print("Done.")

if __name__ == "__main__":
    main()

