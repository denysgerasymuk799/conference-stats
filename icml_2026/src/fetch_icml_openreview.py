"""
Fetch ICML 2026 accepted-paper metadata and save it to data/ICML_2026.xlsx.

The spreadsheet mirrors iclr_2026/data/ICLR_2026.xlsx and adds two columns:
"Spotlight" and "Paper Type".

Two public sources are merged on the OpenReview forum id
-------------------------------------------------------
1. OpenReview  → title, authors, abstract, primary area, keywords, and the
   official venue label ("ICML 2026 regular" / "ICML 2026 spotlight").

   The plain /notes endpoint is behind a Cloudflare Turnstile challenge for
   anonymous clients, and /profiles requires a login, so this script uses the
   /notes/search endpoint, which is still open.  Search needs a term, so it
   sweeps a list of stop-words and unions the results; the sweep stops once
   several consecutive terms stop contributing new papers (in practice "a" +
   "the" + "we" already cover every paper).

2. icml.cc virtual site → per-author institutions, presentation type
   (Oral / Spotlight / Poster), room, poster number, session and times.

   OpenReview exposes no affiliations without a login, so institutions come
   from the conference program instead.  icml.cc publishes it as a static JSON
   snapshot, but the live file is currently truncated to the first API page
   (200 of ~7 000 events) and the paginated API behind it needs an icml.cc
   account, so the script falls back to the most recent *complete* snapshot in
   the Wayback Machine.  Whatever the live file does contain is layered on top.

Paper Type ("Benchmark" / "Evaluation" / "Other") is assigned offline by
classify_paper_type.py from the title and abstract.

Usage:
    python icml_2026/src/fetch_icml_openreview.py [year] [--refresh]
"""

import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

from classify_paper_type import classify as classify_paper_type

DATA_DIR = Path(__file__).parent.parent / "data"
CACHE_DIR = DATA_DIR / "cache"

OPENREVIEW_FORUM = "https://openreview.net/forum?id="
OPENREVIEW_SEARCH = "https://api2.openreview.net/notes/search"
ICML_STATIC = "https://icml.cc/static/virtual/data/icml-{year}-orals-posters.json"
WAYBACK_CDX = "http://web.archive.org/cdx/search/cdx"
WAYBACK_RAW = "https://web.archive.org/web/{ts}id_/{url}"

USER_AGENT = "conference-stats/1.0 (+https://github.com/)"

# Local timezone of the conference venue, used for the schedule columns.
CONFERENCE_TZ = {
    2026: ("Asia/Seoul", "KST"),        # COEX, Seoul, South Korea
    2025: ("America/Vancouver", "PDT"),  # Vancouver, Canada
    2024: ("Europe/Vienna", "CEST"),     # Vienna, Austria
}

# Stop-words swept through OpenReview search.  Ordered most- to least-frequent;
# the sweep exits early once new terms stop adding papers.
SEARCH_TERMS = [
    "a", "the", "we", "and", "of", "to", "in", "for", "that", "is", "on",
    "with", "this", "by", "as", "are", "from", "it", "our", "can", "which",
    "model", "learning", "data", "method", "results", "propose", "show",
]
SEARCH_PAGE = 1000
DRY_TERMS = 3          # stop after this many consecutive terms add nothing

# openpyxl rejects control characters (U+0000–U+001F except \t \n \r, plus U+007F)
_ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
_FORUM_ID = re.compile(r"[?&]id=([A-Za-z0-9_-]+)")


def _clean(s: str) -> str:
    return _ILLEGAL_CHARS.sub("", s).strip()


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def _get(url: str, timeout: int = 120) -> bytes:
    """GET with retries and back-off on rate limiting."""
    delay = 5.0
    last = None
    for attempt in range(6):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except urllib.error.HTTPError as exc:
            last = exc
            if exc.code not in (429, 500, 502, 503, 504):
                raise
        except Exception as exc:            # transient network errors
            last = exc
        time.sleep(delay)
        delay = min(delay * 2, 120)
    raise RuntimeError(f"GET failed after retries: {url} ({last})")


def _get_json(url: str, timeout: int = 120):
    return json.loads(_get(url, timeout))


def _cached(name: str, refresh: bool, produce):
    """Read `name` from data/cache/, otherwise call `produce()` and store it."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path = CACHE_DIR / name
    if path.exists() and not refresh:
        print(f"  (cached) {path.relative_to(DATA_DIR.parent)}")
        return json.loads(path.read_text())
    value = produce()
    path.write_text(json.dumps(value))
    return value


# ---------------------------------------------------------------------------
# Source 1 — OpenReview notes (title / abstract / area / keywords / venue)
# ---------------------------------------------------------------------------

def _search_page(term: str, group: str, offset: int) -> list:
    params = urllib.parse.urlencode({
        "term": term,
        "group": group,
        "source": "forum",
        "limit": SEARCH_PAGE,
        "offset": offset,
    })
    return _get_json(f"{OPENREVIEW_SEARCH}?{params}").get("notes", [])


def fetch_openreview_notes(year: int) -> dict:
    """Return {forum_id: note} for every public note in the conference group."""
    group = f"ICML.cc/{year}/Conference"
    notes: dict[str, dict] = {}
    dry = 0

    for term in SEARCH_TERMS:
        before = len(notes)
        offset = 0
        while True:
            page = _search_page(term, group, offset)
            for note in page:
                notes[note["id"]] = note
            if len(page) < SEARCH_PAGE:
                break
            offset += SEARCH_PAGE
            time.sleep(2)

        added = len(notes) - before
        print(f"  term {term!r:<12} → {len(notes)} unique papers (+{added})")
        dry = dry + 1 if added == 0 else 0
        if dry >= DRY_TERMS:
            break
        time.sleep(2)

    return notes


# ---------------------------------------------------------------------------
# Source 2 — icml.cc program (institutions / schedule / presentation type)
# ---------------------------------------------------------------------------

def _is_complete(payload: dict) -> bool:
    return (
        isinstance(payload, dict)
        and payload.get("next") in (None, "")
        and len(payload.get("results", [])) == payload.get("count")
    )


def _wayback_snapshot(url: str) -> dict | None:
    """Newest archived copy of `url` that is a complete (unpaginated) dump."""
    params = urllib.parse.urlencode({
        "url": url.replace("https://", ""),
        "output": "json",
        "fl": "timestamp,statuscode",
        "filter": "statuscode:200",
    })
    try:
        rows = _get_json(f"{WAYBACK_CDX}?{params}", timeout=60)
    except Exception as exc:
        print(f"  Wayback index unavailable: {exc}")
        return None
    if not rows or len(rows) < 2:
        return None

    for ts, _status in sorted(rows[1:], key=lambda r: r[0], reverse=True):
        print(f"  Trying Wayback snapshot {ts} …")
        try:
            payload = _get_json(WAYBACK_RAW.format(ts=ts, url=url), timeout=300)
        except Exception as exc:
            print(f"    failed: {exc}")
            continue
        if _is_complete(payload):
            print(f"    complete: {payload['count']} events")
            return payload
        print(f"    incomplete ({len(payload.get('results', []))} of "
              f"{payload.get('count')} events), trying older snapshot")
    return None


def fetch_program(year: int) -> list:
    """Return the list of program events published by the icml.cc virtual site."""
    url = ICML_STATIC.format(year=year)

    print(f"  Fetching {url} …")
    live = _get_json(url, timeout=300)
    if _is_complete(live):
        print(f"  → complete: {live['count']} events")
        return live["results"]

    print(f"  → truncated ({len(live.get('results', []))} of {live.get('count')} "
          f"events); the paginated icml.cc API requires an account")
    archived = _wayback_snapshot(url)
    if archived is None:
        print("  WARNING: no complete snapshot found — institutions and schedule "
              "will be missing for most papers")
        return live.get("results", [])

    # Layer the (fresher) live page on top of the archived dump.
    merged = {e["id"]: e for e in archived["results"]}
    merged.update({e["id"]: e for e in live.get("results", [])})
    return list(merged.values())


# ---------------------------------------------------------------------------
# Merge
# ---------------------------------------------------------------------------

_PRESENTATION_RANK = {"Oral": 3, "Spotlight": 2, "Poster": 1}


def _index_program(events: list, year: int) -> dict:
    """
    Group program events by OpenReview forum id.

    Spotlight and oral papers get a talk event *in addition to* their poster
    event, so each paper keeps its best presentation type plus the poster
    event's placement and timing.
    """
    index: dict[str, dict] = {}
    group_url = f"ICML.cc/{year}/Conference"

    for event in events:
        if not (event.get("sourceurl") or "").endswith(group_url):
            continue    # co-located TMLR / JMLR / position-paper talks
        match = _FORUM_ID.search(event.get("paper_url") or "")
        if not match:
            continue

        entry = index.setdefault(match.group(1), {"poster": None, "best": None})
        kind = event.get("eventtype")
        if kind == "Poster" and entry["poster"] is None:
            entry["poster"] = event
        best = entry["best"]
        if best is None or _PRESENTATION_RANK.get(kind, 0) > _PRESENTATION_RANK.get(
                best.get("eventtype"), 0):
            entry["best"] = event

    return index


_AREA_SMALL_WORDS = {"and", "or", "of", "for", "in", "with", "the", "a", "an", "to"}
_AREA_FIXUPS = {
    "Rl": "RL", "Llm": "LLM", "Llms": "LLMs", "Nlp": "NLP",
    "Selfsupervised": "Self-Supervised", "Semisupervised": "Semi-Supervised",
    "Unsupervised": "Unsupervised", "Zeroorder": "Zero-Order",
    "Blackbox": "Black-Box", "Nonconvex": "Non-Convex",
    "Metalearning": "Meta-Learning", "Multitask": "Multi-Task",
    "Multiagent": "Multi-Agent", "Batchoffline": "Batch / Offline",
    "Noregret": "No-Regret", "Highdimensional": "High-Dimensional",
}


def _prettify_area(area: str) -> str:
    """deep_learning->large_language_models → Deep Learning->Large Language Models"""
    out = []
    for part in (area or "").split("->"):
        words = []
        for i, word in enumerate(part.replace("_", " ").split()):
            word = word.title()
            word = _AREA_FIXUPS.get(word, word)
            if i and word.lower() in _AREA_SMALL_WORDS:
                word = word.lower()
            words.append(word)
        if words:
            out.append(" ".join(words))
    return "->".join(out)


def _local_times(event: dict | None, tz: ZoneInfo) -> tuple[str, str, str]:
    if not event or not event.get("starttime"):
        return "", "", ""
    start = datetime.fromisoformat(event["starttime"]).astimezone(tz)
    end_raw = event.get("endtime")
    end = datetime.fromisoformat(end_raw).astimezone(tz) if end_raw else None
    return (
        start.strftime("%Y-%m-%d"),
        start.strftime("%H:%M"),
        end.strftime("%H:%M") if end else "",
    )


def _value(content: dict, key: str, default=""):
    raw = content.get(key)
    if isinstance(raw, dict) and "value" in raw:
        raw = raw["value"]
    return default if raw is None else raw


def build_rows(notes: dict, index: dict, year: int) -> pd.DataFrame:
    tz_name, tz_abbr = CONFERENCE_TZ.get(year, ("UTC", "UTC"))
    tz = ZoneInfo(tz_name)
    columns = _columns(tz_abbr)
    rows = []
    unscheduled = 0

    for forum_id, note in sorted(notes.items()):
        content = note["content"]
        venue = str(_value(content, "venue"))
        if not venue.startswith(f"ICML {year} "):
            continue        # "Submitted to ICML 2026" — public but not accepted

        entry = index.get(forum_id, {})
        best = entry.get("best")
        poster = entry.get("poster") or best

        authors = [str(a) for a in _value(content, "authors", []) or []]
        program_authors = (best or {}).get("authors") or []
        institutions = [str(a.get("institution") or "") for a in program_authors]
        if len(institutions) != len(authors):
            institutions = institutions[:len(authors)]
            institutions += [""] * (len(authors) - len(institutions))

        # The author-declared primary area from OpenReview is authoritative; the
        # icml.cc "topic" field often carries a different, program-assigned area.
        primary_area = _prettify_area(str(_value(content, "primary_area")))
        keywords = "; ".join(str(k) for k in _value(content, "keywords", []) or [])
        title = _clean(str(_value(content, "title")))
        abstract = _clean(str(_value(content, "abstract")))

        date, start, end = _local_times(poster, tz)
        if not date:
            unscheduled += 1

        room = (poster or {}).get("room_name") or ""
        position = (poster or {}).get("poster_position") or ""
        location = " ".join(p for p in (room, position) if p)

        virtual = ((poster or {}).get("virtualsite_url")
                   or (best or {}).get("virtualsite_url") or "")
        spotlight = venue.endswith("spotlight")

        rows.append({
            "Decision":         (best or {}).get("eventtype") or "Poster",
            "Spotlight":        "Yes" if spotlight else "No",
            "Title":            title,
            "Authors":          "; ".join(authors),
            "Institutions":     "; ".join(institutions),
            "Abstract":         abstract,
            "Primary Area":     _clean(str(primary_area)),
            "Keywords":         _clean(keywords),
            "Paper Type":       classify_paper_type(title, abstract, primary_area, keywords),
            f"Date ({tz_abbr})":  date,
            f"Start ({tz_abbr})": start,
            f"End ({tz_abbr})":   end,
            "Location":         location,
            "Poster #":         position,
            "Session":          (poster or {}).get("session") or "",
            "OpenReview URL":   OPENREVIEW_FORUM + forum_id,
            "Virtual Site URL": f"https://icml.cc{virtual}" if virtual else "",
        })

    if unscheduled:
        print(f"  {unscheduled} accepted papers have no in-person session "
              f"(schedule columns left blank)")
    return pd.DataFrame(rows, columns=columns)


def _columns(tz_abbr: str) -> list:
    return [
        "Decision", "Spotlight", "Title", "Authors", "Institutions", "Abstract",
        "Primary Area", "Keywords", "Paper Type",
        f"Date ({tz_abbr})", f"Start ({tz_abbr})", f"End ({tz_abbr})",
        "Location", "Poster #", "Session",
        "OpenReview URL", "Virtual Site URL",
    ]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def save(year: int, df: pd.DataFrame):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / f"ICML_{year}.xlsx"
    df.to_excel(path, index=False, sheet_name="All Accepted")
    print(f"  Saved → {path}  ({len(df)} rows)")


def summarise(df: pd.DataFrame):
    print(f"\n  Accepted papers        {len(df)}")
    for label, count in df["Decision"].value_counts().items():
        print(f"    {label:<20} {count}")
    print(f"    {'spotlighted':<20} {(df['Spotlight'] == 'Yes').sum()}")
    print("  Paper type")
    for label, count in df["Paper Type"].value_counts().items():
        print(f"    {label:<20} {count}  ({count / len(df):.1%})")
    missing = (df["Institutions"].str.strip() == "").sum()
    print(f"  Papers with no institutions   {missing}")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    refresh = "--refresh" in sys.argv
    year = int(args[0]) if args else 2026

    print(f"\n{'=' * 60}\n  ICML {year}\n{'=' * 60}")

    print("\nOpenReview (titles, abstracts, areas, keywords, venue)")
    notes = _cached(f"openreview_{year}.json", refresh,
                    lambda: fetch_openreview_notes(year))
    print(f"  {len(notes)} public notes in ICML.cc/{year}/Conference")

    print("\nicml.cc program (institutions, schedule, presentation type)")
    events = _cached(f"program_{year}.json", refresh, lambda: fetch_program(year))
    index = _index_program(events, year)
    print(f"  {len(events)} events → {len(index)} conference-track papers")

    print("\nMerging")
    df = build_rows(notes, index, year)
    if df.empty:
        print("  ERROR: no accepted papers found")
        sys.exit(1)

    save(year, df)
    summarise(df)


if __name__ == "__main__":
    main()
