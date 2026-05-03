"""
Fetch ICLR paper metadata from OpenReview for years 2022–2025.
Saves one .xlsx file per year into data/.

API details per year:
  2022, 2023 → OpenReview API v1  (api.openreview.net)   – plain content values
  2024, 2025 → OpenReview API v2  (api2.openreview.net)  – content values wrapped in {"value": ...}

Acceptance is detected from the `venue` field (present in all years):
  e.g. "ICLR 2025 Poster", "ICLR 2023 notable top 5%", "ICLR 2024 oral"

Output columns mirror data/ICLR_2026.xlsx:
  Decision, Title, Authors, Institutions, Abstract, Primary Area, Keywords,
  Date (BRT), Start (BRT), End (BRT), Location, Poster #, Session,
  OpenReview URL, Virtual Site URL
"""

import re
import sys
import time
import traceback
from pathlib import Path

import openreview
import pandas as pd

DATA_DIR = Path(__file__).parent.parent / "data"
DATA_DIR.mkdir(exist_ok=True)

OPENREVIEW_FORUM = "https://openreview.net/forum?id="

# openpyxl rejects control characters (U+0000–U+001F except \t \n \r, plus U+007F)
_ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def _clean(s: str) -> str:
    return _ILLEGAL_CHARS.sub("", s)


# ---------------------------------------------------------------------------
# Content-value helpers (handle v1 plain vs v2 wrapped {"value": ...})
# ---------------------------------------------------------------------------

def _val(raw):
    """Unwrap {"value": x} or return raw."""
    if isinstance(raw, dict) and "value" in raw:
        return raw["value"]
    return raw


def _str_val(content: dict, *keys) -> str:
    for k in keys:
        if k in content:
            v = _val(content[k])
            if v is None:
                continue
            if isinstance(v, list):
                return _clean("; ".join(str(x) for x in v if x))
            return _clean(str(v))
    return ""


def _list_val(content: dict, *keys) -> list:
    for k in keys:
        if k in content:
            v = _val(content[k])
            if isinstance(v, list):
                return v
            if v:
                return [v]
    return []


# ---------------------------------------------------------------------------
# Decision parsing
# ---------------------------------------------------------------------------

def parse_decision(venue_raw) -> str | None:
    """
    Return a normalised decision label or None if the paper is not accepted.
    Works for all years:
      "ICLR 2022 Poster/Spotlight/Oral"
      "ICLR 2023 poster / notable top 5% / notable top 25%"
      "ICLR 2024 poster/oral/spotlight"
      "ICLR 2025 Poster/Spotlight/Oral"
    """
    v = str(_val(venue_raw) if isinstance(venue_raw, dict) else venue_raw).lower()

    if not v or "submitted" in v or "withdrawn" in v or "rejected" in v or "desk" in v:
        return None

    if "notable top 5" in v:
        return "Notable Top 5%"
    if "notable top 25" in v:
        return "Notable Top 25%"
    if "oral" in v:
        return "Oral"
    if "spotlight" in v:
        return "Spotlight"
    if "poster" in v:
        return "Poster"
    if "accept" in v:
        return "Accept"

    return None


# ---------------------------------------------------------------------------
# Single-note extraction
# ---------------------------------------------------------------------------

def _extract_row(note, decision: str, year: int, is_v2: bool) -> dict:
    c = note.content

    authors_list = _list_val(c, "authors")
    institutions_list = _list_val(c, "institutions", "affiliations")

    # Primary area: field name varies by year
    primary_area = _str_val(
        c,
        "primary_area",                                          # 2024, 2025
        "Please_choose_the_closest_area_that_your_submission_falls_into",  # 2023
        "subject_areas",                                         # some older years
        "track",
    )

    keywords = _str_val(c, "keywords")
    abstract = _str_val(c, "abstract")
    title = _str_val(c, "title")

    forum_id = note.id if is_v2 else (getattr(note, "forum", None) or note.id)
    openreview_url = OPENREVIEW_FORUM + forum_id
    virtual_url = f"https://iclr.cc/virtual/{year}/poster/{forum_id}"

    return {
        "Decision":         decision,
        "Title":            title,
        "Authors":          "; ".join(str(a) for a in authors_list) if authors_list else "",
        "Institutions":     "; ".join(str(i) for i in institutions_list) if institutions_list else "",
        "Abstract":         abstract,
        "Primary Area":     primary_area,
        "Keywords":         keywords,
        "Date (BRT)":       "",
        "Start (BRT)":      "",
        "End (BRT)":        "",
        "Location":         "",
        "Poster #":         "",
        "Session":          "",
        "OpenReview URL":   openreview_url,
        "Virtual Site URL": virtual_url,
    }


def _columns():
    return [
        "Decision", "Title", "Authors", "Institutions", "Abstract",
        "Primary Area", "Keywords",
        "Date (BRT)", "Start (BRT)", "End (BRT)",
        "Location", "Poster #", "Session",
        "OpenReview URL", "Virtual Site URL",
    ]


# ---------------------------------------------------------------------------
# Per-year fetch
# ---------------------------------------------------------------------------

def _fetch(year: int) -> pd.DataFrame:
    is_v2 = year >= 2024

    if is_v2:
        print(f"  Using OpenReview API v2")
        client = openreview.api.OpenReviewClient(baseurl="https://api2.openreview.net")
        invitation = f"ICLR.cc/{year}/Conference/-/Submission"
    else:
        print(f"  Using OpenReview API v1")
        client = openreview.Client(baseurl="https://api.openreview.net")
        invitation = f"ICLR.cc/{year}/Conference/-/Blind_Submission"

    print(f"  Fetching all submissions ({invitation}) …")
    notes = client.get_all_notes(invitation=invitation)
    print(f"  → {len(notes)} total submissions")

    rows = []
    skipped = 0
    for note in notes:
        venue_raw = note.content.get("venue", "")
        decision = parse_decision(venue_raw)
        if decision is None:
            skipped += 1
            continue
        rows.append(_extract_row(note, decision, year, is_v2))

    print(f"  → {len(rows)} accepted  |  {skipped} rejected/withdrawn/pending")
    return pd.DataFrame(rows, columns=_columns())


# ---------------------------------------------------------------------------
# Save
# ---------------------------------------------------------------------------

def save(year: int, df: pd.DataFrame):
    path = DATA_DIR / f"ICLR_{year}.xlsx"
    df.to_excel(path, index=False, sheet_name="All Accepted")
    print(f"  Saved → {path}  ({len(df)} rows)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

YEARS = [2025, 2024, 2023, 2022]


def main():
    target_years = YEARS
    if len(sys.argv) > 1:
        try:
            target_years = [int(y) for y in sys.argv[1:]]
        except ValueError:
            print(f"Usage: python fetch_iclr_openreview.py [year ...]")
            sys.exit(1)

    for year in target_years:
        print(f"\n{'='*60}")
        print(f"  ICLR {year}")
        print(f"{'='*60}")
        try:
            df = _fetch(year)
            if df.empty:
                print(f"  WARNING: no accepted papers found for {year}")
            else:
                save(year, df)
        except Exception:
            print(f"  ERROR fetching {year}:")
            traceback.print_exc()
        time.sleep(1)

    print("\nDone.")


if __name__ == "__main__":
    main()
