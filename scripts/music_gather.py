#!/usr/bin/env python3
"""
music_gather.py – Weekly Automation music retrieval module

Responsibilities
----------------
• Fill in song text/XML placeholders in the Master markdown file:
    - Opening songs: use CCLI plaintext lyrics (from config.ccli_dir)
      {song_opening_title_elkton}  →  {song_opening_text_elkton}
      {song_opening_title_lb}      →  {song_opening_text_lb}

    - Middle songs: use OpenLP songs DB (via utils.openlp)
      {song_middle_title_elkton}
      {song_middle_id_elkton}      →  {song_middle_text_elkton}
      (and LB equivalents)

    - Closing songs: same as middle songs, with LB special-case
      "Congregational choice" (skip lookup for LB).

• All paths come from weekly_config.ini via utils.config.Config.
• All file IO via utils.file_io.
• Placeholder handling via utils.placeholder.
• Text cleaning via utils.text_clean.
• OpenLP access via utils.openlp (build_song_index, load_song).
"""

from __future__ import annotations

import argparse
import logging
import re
from dataclasses import dataclass

from scripts.utils.config import config, resolve_master_path
from scripts.utils.file_io import read_text, write_text
from scripts.utils.placeholder import (
    append_below_placeholder,
    extract_block,
)
from scripts.utils.text_clean import clean_markdown, clean_text, xml_to_text
from scripts.utils.openlp import build_song_index, load_song, song_xml_to_labeled_markdown



# =====================================================================
# Standardized diagnostic insertion
# =====================================================================

def append_missing(md: str, placeholder: str, title: str, church: str, reason: str) -> str:
    """
    Standardized diagnostic insertion for missing songs.
    Example message:
        [Missing song for elkton: Amazing Grace — invalid song ID]
    """
    msg = f"[Missing song for {church}: {title} — {reason}]"
    return append_below_placeholder(md, placeholder, msg)


# =====================================================================
# CCLI plaintext lyrics (from config.ccli_dir)
# =====================================================================

def get_ccli_lyrics(title: str) -> str | None:
    """
    Fetch CCLI lyrics from a folder structure like:
        {ccli_dir}/{title}/{title} lyrics.txt

    The returned lyrics are lightly cleaned for markdown safety.
    """
    title = clean_text(title or "").strip()
    if not title:
        return None

    folder = config.ccli_dir / title
    file_path = folder / f"{title} lyrics.txt"
    if not file_path.exists():
        logging.warning("CCLI lyrics file not found: %s", file_path)
        return None

    raw = read_text(file_path).strip()
    if not raw:
        return None

    return clean_markdown(raw)


# =====================================================================
# Hymnal mapping
# =====================================================================

def hymnal_from_prefix(church: str, prefix: str) -> str | None:
    """
    Map the R/B/K prefix in IDs to the proper OpenLP hymnal name
    for each church.
    """
    prefix = prefix.lower()
    church = church.lower()

    if church == "elkton":
        match prefix:
            case "r":
                return "United Methodist Hymnal"
            case "b":
                return "The Singing Church"


    elif church == "lb":
        match prefix:
            case "r":
                return "United Methodist Hymnal"
            case "b":
                return "Hymns of Faith"
            case "k":
                return "The Faith We Sing"

    logging.error("Unknown hymnal prefix %r for %s", prefix, church)
    return None


# =====================================================================
# Song ID parsing
# =====================================================================

@dataclass
class ParsedSongID:
    prefix: str              # r / b / k
    entry: str               # numeric hymnal entry, as string
    verses: list[int] | None # optional requested verses list


def parse_song_id(raw: str | None, title_for_log: str = "") -> ParsedSongID | None:
    """
    Parse IDs like:
        'r 123'
        'B218 vv 1-3,5'
        'k 42 vv 1, 3, 5-7'

    Returns ParsedSongID or None on failure.
    """
    if raw is None:
        return None

    # Strip comment-style suffix (%% …)
    base = raw.split("%%", 1)[0]
    cleaned = clean_text(base)
    if not cleaned:
        return None

    pattern = re.compile(
        r"([rbk])\s*(\d+)(?:\s+v+v?\s+(.+))?$",
        re.IGNORECASE,
    )
    m = pattern.match(cleaned)
    if not m:
        logging.warning("Could not parse song id %r for title %r", raw, title_for_log)
        return None

    prefix = m.group(1).lower()
    entry_str = m.group(2)
    verses_raw = m.group(3)

    verses: list[int] | None = None
    if verses_raw:
        nums = re.findall(r"\d+", verses_raw)
        if nums:
            verses = [int(n) for n in nums]

    return ParsedSongID(prefix=prefix, entry=entry_str, verses=verses)


# =====================================================================
# XML → plain text
# =====================================================================



def resolve_openlp_song(
    church: str,
    index: dict,
    title: str,
    id_raw: str | None,
):
    """
    Given:
        church: 'elkton' or 'lb'
        index:  (hymnal, entry) → uuid (from build_song_index(church))
        title:  song title from Master (for diagnostics)
        id_raw: like 'r 123 vv 1-3'

    Return:
        text string

    On error:
        error_message_string starting with one of:
            "No valid song id"
            "Invalid hymnal prefix"
            "Hymn not found"
            "Lyrics not found"
    """
    if not title and not id_raw:
        return "No valid song id for title ‘’"

    parsed = parse_song_id(id_raw, title_for_log=title)
    if not parsed:
        return f"No valid song id for title {title!r}"

    hymnal = hymnal_from_prefix(church, parsed.prefix)
    if not hymnal:
        return f"Invalid hymnal prefix for {title!r}"

    key = (hymnal, str(parsed.entry))
    uuid = index.get(key)
    if not uuid:
        return f"Hymn not found in {hymnal}: {title}"

    song = load_song(church, int(uuid))
    if not song:
        return f"Lyrics not found for UUID {uuid}"

    lyrics_xml = str(song["lyrics"] or "").strip()
    if not lyrics_xml:
        return f"Lyrics not found for UUID {uuid}"

    text = song_xml_to_labeled_markdown(lyrics_xml, verses_filter=parsed.verses)
    if text is None:
        text = xml_to_text(lyrics_xml)
        if parsed.verses:
            verse_note = "Verses: " + ", ".join(map(str, parsed.verses))
            text = f"{verse_note}\n\n{text}"

    return text.strip()


# =====================================================================
# Processing the Master file
# =====================================================================

_OPENLP_ERROR_REASONS = {
    "No valid song id":      "invalid song ID",
    "Invalid hymnal prefix": "invalid hymnal prefix",
    "Hymn not found":        "not in hymnal index",
    "Lyrics not found":      "missing lyrics XML",
}


def _gather_song_text(church: str, index: dict, title: str, id_raw: str | None) -> str:
    """
    Route a single song slot to hymnal or praise-song lookup.

    If id_raw parses as a hymnal reference (r/b/k + digits), look it up in
    the OpenLP DB and return labeled markdown.  Otherwise treat the slot as a
    praise song and fetch CCLI plain-text lyrics by title.

    Always returns a non-empty string; error strings begin with "[Missing".
    """
    if parse_song_id(id_raw or "", title_for_log=title):
        result = resolve_openlp_song(church, index, title, id_raw)
        for prefix, reason in _OPENLP_ERROR_REASONS.items():
            if result.startswith(prefix):
                return f"[Missing song for {church}: {title} — {reason}]"
        return result

    # Praise song: look up CCLI plain-text lyrics by title.
    lyr = get_ccli_lyrics(clean_text(title))
    if lyr:
        return lyr
    return f"[Missing song for {church}: {title} — CCLI lyrics not found]"


def process_master(md: str, elk_index: dict, lb_index: dict) -> str:
    """
    Fill in song text in the Master markdown for all three song slots.

    Each slot is routed automatically: a hymnal ID (r/b/k + digits) triggers
    the OpenLP DB path with labeled verse/chorus headings; anything else
    (artist name, blank) triggers the CCLI plain-text path.
    """
    for slot in ("opening", "middle", "closing"):
        for church, index in (("elkton", elk_index), ("lb", lb_index)):
            title_key = f"song_{slot}_title_{church}"
            id_key    = f"song_{slot}_id_{church}"
            text_key  = f"song_{slot}_text_{church}"

            title = extract_block(md, title_key)
            sid   = extract_block(md, id_key)

            if not title and not sid:
                continue

            # LB closing special case: congregational choice → skip lookup.
            if slot == "closing" and church == "lb" and title.lower() == "congregational choice":
                continue

            text = _gather_song_text(church, index, title, sid)
            if text:
                md = append_below_placeholder(md, text_key, text)

    return md



# =====================================================================
# Main
# =====================================================================

def main():
    parser = argparse.ArgumentParser(description="Weekly Automation: music retrieval")
    parser.add_argument("--master", required=True, help="Master markdown filename or path")
    parser.add_argument("--verbose", action="store_true", help="Enable debug logging")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    master_path = resolve_master_path(args.master)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    # Build song indexes once per church from OpenLP DB
    logging.info("Building OpenLP song indexes...")
    elk_index = build_song_index("elkton")
    lb_index = build_song_index("lb")

    md = read_text(master_path)
    updated = process_master(md, elk_index, lb_index)

    if updated != md:
        write_text(master_path, updated)
        logging.info("Updated: %s", master_path)
    else:
        logging.info("No changes made.")


if __name__ == "__main__":
    main()
