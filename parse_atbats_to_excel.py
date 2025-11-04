#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
parse_atbats_to_excel.py
Parses ESPN MLB Play-by-Play pages (live or saved HTML) and writes:
  - AtBats sheet (one row per at-bat, with inning/half, description, pitch summary, and PITCHER)
  - Pitches sheet (one row per pitch, with result/type/MPH and PITCHER)
  - PitchingChanges sheet (detected "Pitching Change: A replaces B." events)
Supports JS-rendered pages via Playwright (--render).
"""

import argparse
import re
import sys
import time
from dataclasses import dataclass, asdict
from typing import List, Tuple, Dict, Optional

from bs4 import BeautifulSoup
from bs4.element import Tag

# Optional imports resolved at runtime
try:
    import requests
except Exception:
    requests = None

import pandas as pd


# ---------------------------
# Utilities
# ---------------------------

UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

PITCH_CHANGE_RX = re.compile(
    r"(?:Pitching\s*Change|Pitching change)\s*:?\s*(?P<new>[^.,]+?)\s+replaces\s+(?P<old>[^.,]+)",
    flags=re.IGNORECASE,
)

PITCHING_INLINE_RX = re.compile(
    r"Pitching\s*:?\s*(?P<pitcher>[A-Za-z\.\-’' ]+)", flags=re.IGNORECASE
)

INNING_HEADER_RX = re.compile(
    r"(Top|Bottom)\s+(\d+)(?:st|nd|rd|th)", flags=re.IGNORECASE
)


@dataclass
class AtBatRow:
    inning: int
    half: str
    team: str
    description: str
    pitch_count: int
    pitch_sequence: str
    header_id: str
    body_id: str
    pitcher: str


@dataclass
class PitchRow:
    inning: int
    half: str
    team: str
    description: str
    pitch_no: Optional[int]
    pitch_result: str
    pitch_type: str
    mph: Optional[int]
    header_id: str
    body_id: str
    pitcher: str


@dataclass
class PitchChangeRow:
    inning: int
    half: str
    new_pitcher: str
    old_pitcher: str
    description: str


# ---------------------------
# Fetching / Rendering
# ---------------------------

def render_with_playwright(url: str, quiet: bool = False, timeout_ms: int = 25000) -> str:
    """
    Load URL with Playwright (Chromium), expand accordions, and return page HTML.
    Requires: pip install playwright && playwright install chromium
    """
    try:
        from playwright.sync_api import sync_playwright
    except Exception as e:
        raise RuntimeError(
            "Playwright not available. Install with:\n"
            "  pip install playwright\n  playwright install chromium"
        ) from e

    if not quiet:
        print(f"Rendering URL (Playwright): {url}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(user_agent=UA, viewport={"width": 1400, "height": 2000})
        page = ctx.new_page()
        page.goto(url, timeout=timeout_ms)
        # Wait for core container
        page.wait_for_timeout(1500)

        # Try to expand all at-bat accordions
        # Buttons have class 'AtBatAccordion__header'
        try:
            buttons = page.query_selector_all("button.AtBatAccordion__header")
            for b in buttons:
                aria_expanded = (b.get_attribute("aria-expanded") or "").lower()
                if aria_expanded != "true":
                    b.click()
            page.wait_for_timeout(800)
        except Exception:
            pass

        html = page.content()
        browser.close()
        return html


def fetch_html(src: str, quiet: bool = False) -> str:
    """
    If src looks like a URL, fetch via requests (with UA). Otherwise read local file.
    """
    if src.startswith("http://") or src.startswith("https://"):
        if requests is None:
            raise RuntimeError("requests not available; cannot fetch URL. `pip install requests`")
        if not quiet:
            print(f"Fetching URL: {src}")
        resp = requests.get(src, headers={"User-Agent": UA})
        try:
            resp.raise_for_status()
        except Exception as e:
            # ESPN may 403 without JS. Suggest using --render.
            if resp.status_code == 403:
                raise RuntimeError(
                    "403 Forbidden fetching ESPN. Try using --render to let Playwright load JS."
                ) from e
            raise
        return resp.text
    else:
        if not quiet:
            print(f"Reading file: {src}")
        with open(src, "r", encoding="utf-8") as f:
            return f.read()


# ---------------------------
# Parsing helpers
# ---------------------------

def text_or_empty(node: Optional[Tag]) -> str:
    return node.get_text(" ", strip=True) if node else ""


def get_inning_half_from_context(atbat_header: Tag) -> Tuple[int, str, str]:
    """
    Best-effort extraction of (inning_num, half, team_name) from nearby DOM.
    Team is often not available inline; we return 'Unknown Team' if not found.
    """
    # Walk up to a reasonable container, then search backward for a header with inning info.
    inning_num, half, team = 0, "Unknown", "Unknown Team"

    # Try immediate parent/siblings for an inning label
    cursor = atbat_header
    for _ in range(8):  # bubble up a few levels
        if not cursor or not isinstance(cursor, Tag):
            break
        # Look among previous siblings for recognizable headers
        sib = cursor.previous_sibling
        # Iterate backwards through siblings
        while sib:
            if isinstance(sib, Tag):
                text = sib.get_text(" ", strip=True)
                m = INNING_HEADER_RX.search(text or "")
                if m:
                    half = m.group(1).title()
                    inning_num = int(m.group(2))
                    # Team name is usually not in that header; keep Unknown Team
                    return inning_num, half, team
            sib = sib.previous_sibling
        cursor = cursor.parent

    # Fallback defaults
    return inning_num, half, team


def parse_pitch_rows_from_body(body_div: Tag) -> List[Tuple[Optional[int], str, str, Optional[int]]]:
    """
    From a .AtBatAccordion__body div, parse the inner PitchTable rows.
    Returns list of (pitch_no, pitch_result, pitch_type, mph).
    """
    out = []
    if not body_div:
        return out

    # Find the inner table
    table = body_div.select_one("div.Table__Scroller table.Table")
    if not table:
        return out

    tbody = table.find("tbody")
    if not tbody:
        return out

    for tr in tbody.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 3:
            continue

        # Col 0: has PitchCountIcon and a span with text like "Ball", "Foul Ball", "Single", etc.
        pitch_no = None
        pitch_result = ""
        try:
            icon = tds[0].select_one(".PitchCountIcon")
            if icon and icon.get_text(strip=True).isdigit():
                pitch_no = int(icon.get_text(strip=True))
        except Exception:
            pass

        # Result text (span next to icon)
        result_span = tds[0].find("span")
        if result_span:
            pitch_result = result_span.get_text(strip=True)

        # Col 1: pitch type
        pitch_type = text_or_empty(tds[1])

        # Col 2: mph (may be blank)
        mph_txt = text_or_empty(tds[2])
        mph_val = None
        if mph_txt:
            try:
                mph_val = int(mph_txt)
            except Exception:
                mph_val = None

        out.append((pitch_no, pitch_result, pitch_type, mph_val))

    return out


def detect_pitching_change(desc_text: str) -> Optional[Tuple[str, str]]:
    """
    If description encodes a pitching change, return (new_pitcher, old_pitcher), else None.
    """
    m = PITCH_CHANGE_RX.search(desc_text)
    if m:
        new_p = m.group("new").strip()
        old_p = m.group("old").strip()
        return new_p, old_p
    return None


def detect_inline_pitcher(desc_text: str) -> Optional[str]:
    """
    Some sites include "Pitching: Name" inline. Try to pick it up.
    """
    m = PITCHING_INLINE_RX.search(desc_text)
    if m:
        return m.group("pitcher").strip()
    return None


# ---------------------------
# Main parsing routine
# ---------------------------

def parse_play_by_play(html: str, verbose: bool = True) -> Tuple[List[AtBatRow], List[PitchRow], List[PitchChangeRow]]:
    soup = BeautifulSoup(html, "lxml")

    atbats: List[AtBatRow] = []
    pitches: List[PitchRow] = []
    changes: List[PitchChangeRow] = []

    # Track current pitcher per half-inning: key = (inning, half)
    current_pitcher: Dict[Tuple[int, str], str] = {}

    # Find all at-bat blocks
    atbat_blocks = soup.select("div.AtBatAccordion")
    if verbose:
        print(f"Found {len(atbat_blocks)} at-bats on the page.")

    for idx, block in enumerate(atbat_blocks):
        header = block.select_one("button.AtBatAccordion__header")
        if not header:
            continue

        # Description (what happened in the play)
        desc = text_or_empty(header.select_one(".PlayHeader__description"))

        # aria-controls gives us the body id containing the pitch table
        body_id = header.get("aria-controls") or ""
        body = None
        if body_id:
            body = block.find(id=body_id)

        # Basic inning/half/team inference (best-effort)
        inning, half, team = get_inning_half_from_context(header)

        # Try to identify pitcher:
        # 1) If description encodes a "Pitching Change", update state
        # 2) Else if description includes "Pitching: X", seed the state
        key = (inning, half)
        new_pitcher_inline = detect_inline_pitcher(desc)
        pc = detect_pitching_change(desc)
        if pc:
            new_p, old_p = pc
            # Update the state immediately
            current_pitcher[key] = new_p
            changes.append(PitchChangeRow(inning=inning or 0, half=half, new_pitcher=new_p, old_pitcher=old_p, description=desc))
            if verbose:
                print(f"   • Pitching change detected: {old_p} → {new_p}")
        elif new_pitcher_inline:
            current_pitcher[key] = new_pitcher_inline
            if verbose:
                print(f"   • Detected pitcher inline: {new_pitcher_inline}")

        # Current pitcher for this at-bat (may be empty if none detected yet for this half)
        pitcher_for_ab = current_pitcher.get(key, "")

        # Parse pitch table (if present)
        pitch_rows = parse_pitch_rows_from_body(body) if body else []

        # Build pitch summary string like "P1: Ball, P2: Ball, ..."
        if pitch_rows:
            seq_parts = []
            for (pno, pres, ptype, mph) in pitch_rows:
                label = f"P{pno}" if pno is not None else "P?"
                seq_parts.append(f"{label}: {pres}")
            pitch_summary = " • " + ", ".join(seq_parts)
        else:
            pitch_summary = ""

        # Status output
        if verbose:
            inn_show = inning if inning else "Unknown Inning"
            print(f"Parsing Inning {inn_show} ({half} – {team}): {desc}")
            if pitch_rows:
                joined = ", ".join([f"P{p[0] if p[0] is not None else '?'}: {p[1]}" for p in pitch_rows])
                print(f"   • Captured {len(pitch_rows)} pitch rows → {joined}")
            else:
                print(f"   • Captured 0 pitch rows")

        # Create AtBatRow
        atbats.append(
            AtBatRow(
                inning=inning or 0,
                half=half,
                team=team,
                description=desc,
                pitch_count=len(pitch_rows),
                pitch_sequence=pitch_summary.replace(" • ", ""),
                header_id=header.get("id") or "",
                body_id=body_id,
                pitcher=pitcher_for_ab,
            )
        )

        # Create PitchRow entries
        for (pno, pres, ptype, mph) in pitch_rows:
            pitches.append(
                PitchRow(
                    inning=inning or 0,
                    half=half,
                    team=team,
                    description=desc,
                    pitch_no=pno,
                    pitch_result=pres,
                    pitch_type=ptype,
                    mph=mph,
                    header_id=header.get("id") or "",
                    body_id=body_id,
                    pitcher=pitcher_for_ab,
                )
            )

        # Some sites encode pitching change as its own at-bat with no table; handled above.
        # Also: if we encounter explicit "Pitching: X" with no change, we already seeded state.

    return atbats, pitches, changes


# ---------------------------
# CLI / Excel output
# ---------------------------

def main():
    ap = argparse.ArgumentParser(description="Parse ESPN Play-by-Play at-bats & pitches → Excel.")
    ap.add_argument("-i", "--input", help="URL or local HTML file", required=True)
    ap.add_argument("-o", "--output", help="Output .xlsx path", default="out.xlsx")
    ap.add_argument("--render", action="store_true", help="Use Playwright to render (JS) before parsing")
    ap.add_argument("--save-html", help="If set, save the rendered/fetched HTML to this path")
    ap.add_argument("--quiet", action="store_true", help="Minimal console output")

    args = ap.parse_args()

    # Fetch/render
    if args.render:
        html = render_with_playwright(args.input, quiet=args.quiet)
    else:
        html = fetch_html(args.input, quiet=args.quiet)

    # Optionally save the HTML we actually parsed
    if args.save_html:
        with open(args.save_html, "w", encoding="utf-8") as f:
            f.write(html)
        if not args.quiet:
            print(f"Saved HTML to: {args.save_html}")

    # Parse
    atbats, pitch_rows, changes = parse_play_by_play(html, verbose=(not args.quiet))

    # DataFrames
    df_ab = pd.DataFrame([asdict(x) for x in atbats])
    df_p = pd.DataFrame([asdict(x) for x in pitch_rows])
    df_c = pd.DataFrame([asdict(x) for x in changes])

    # Write Excel
    out_path = args.output
    if not out_path.lower().endswith(".xlsx"):
        out_path += ".xlsx"

    if not args.quiet:
        print(f"Writing Excel: {out_path}")

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        # Keep friendly column order
        if not df_ab.empty:
            df_ab = df_ab[
                ["inning", "half", "team", "pitcher", "description", "pitch_count", "pitch_sequence", "header_id", "body_id"]
            ]
        if not df_p.empty:
            df_p = df_p[
                ["inning", "half", "team", "pitcher", "description", "pitch_no", "pitch_result", "pitch_type", "mph", "header_id", "body_id"]
            ]
        if not df_c.empty:
            df_c = df_c[["inning", "half", "new_pitcher", "old_pitcher", "description"]]

        # Sort for readability
        if not df_ab.empty:
            df_ab.sort_values(by=["inning", "half"], inplace=True, ignore_index=True)
        if not df_p.empty:
            df_p.sort_values(by=["inning", "half", "pitch_no"], inplace=True, ignore_index=True)
        if not df_c.empty:
            df_c.sort_values(by=["inning", "half"], inplace=True, ignore_index=True)

        (df_ab if not df_ab.empty else pd.DataFrame()).to_excel(xw, sheet_name="AtBats", index=False)
        (df_p if not df_p.empty else pd.DataFrame()).to_excel(xw, sheet_name="Pitches", index=False)
        (df_c if not df_c.empty else pd.DataFrame()).to_excel(xw, sheet_name="PitchingChanges", index=False)

    if not args.quiet:
        print("Done.")


if __name__ == "__main__":
    main()

