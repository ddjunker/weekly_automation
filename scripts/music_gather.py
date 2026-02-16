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
      {song_middle_id_elkton}      →  {song_middle_text_elkton}, {song_middle_xml_elkton}
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
from pathlib import Path

import yaml

from scripts.utils.config import config
from scripts.utils.file_io import read_text, write_text
from scripts.utils.placeholder import append_below_placeholder, extract_block
from scripts.utils.text_clean import clean_markdown, clean_text, xml_to_text
from scripts.utils.openlp import build_song_index, load_song


# =====================================================================
# Master path resolution (mirrors text_gather.py)
# =====================================================================

def resolve_master_path(master_arg: str) -> Path:
    """
    Resolve the Master markdown file path:

    1. If master_arg is an absolute path → return it unchanged.
    2. Otherwise → return config.worship_dir / master_arg
    3. worship_dir must exist (strict behavior).
    """
    mp = Path(master_arg)

    # Case 1: absolute override
    if mp.is_absolute():
        return mp

    # Case 2: relative → use worship_dir
    worship = config.worship_dir.expanduser().resolve()
    if not worship.exists():
        raise FileNotFoundError(f"Worship directory not found: {worship}")

    return worship / master_arg


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
        (xml, text)

    On error:
        (None, error_message_string)
        where the error string starts with one of:
            "No valid song id"
            "Invalid hymnal prefix"
            "Hymn not found"
            "Lyrics not found"
    """
    if not title and not id_raw:
        return None, "No valid song id for title ''"

    parsed = parse_song_id(id_raw, title_for_log=title)
    if not parsed:
        return None, f"No valid song id for title {title!r}"

    hymnal = hymnal_from_prefix(church, parsed.prefix)
    if not hymnal:
        return None, f"Invalid hymnal prefix for {title!r}"

    key = (hymnal, str(parsed.entry))
    uuid = index.get(key)
    if not uuid:
        return None, f"Hymn not found in {hymnal}: {title}"

    song = load_song(church, int(uuid))
    if not song:
        return None, f"Lyrics not found for UUID {uuid}"

    lyrics_xml = str(song["lyrics"] or "").strip()
    if not lyrics_xml:
        return None, f"Lyrics not found for UUID {uuid}"

    text = xml_to_text(lyrics_xml)

    # We can’t surgically subset verses from XML here, so just annotate.
    if parsed.verses:
        verse_note = "Verses: " + ", ".join(map(str, parsed.verses))
        text = f"{verse_note}\n\n{text}"

    return lyrics_xml, text.strip()


# =====================================================================
# Processing the Master file
# =====================================================================

def process_master(md: str, elk_index: dict, lb_index: dict) -> str:
    """
    Fill in song text/XML in the Master markdown using:
      • CCLI plaintext for opening songs
      • OpenLP DB + index for middle and closing songs
    """

    # ----------------------
    # Opening songs (CCLI)
    # ----------------------
    for church in ("elkton", "lb"):
        title_key = f"song_opening_title_{church}"
        text_key = f"song_opening_text_{church}"

        title = extract_block(md, title_key)
        title_clean = clean_text(title)

        if title_clean:
            lyr = get_ccli_lyrics(title_clean)
            if lyr:
                md = append_below_placeholder(md, text_key, lyr)
            else:
                md = append_missing(
                    md,
                    text_key,
                    title_clean,
                    church,
                    "CCLI lyrics not found",
                )

    # ----------------------
    # Middle songs (OpenLP)
    # ----------------------
    for church, cmap in (("elkton", elk_index), ("lb", lb_index)):
        title_key = f"song_middle_title_{church}"
        id_key = f"song_middle_id_{church}"
        text_key = f"song_middle_text_{church}"
        xml_key = f"song_middle_xml_{church}"

        title = extract_block(md, title_key)
        sid = extract_block(md, id_key)

        if not title and not sid:
            continue

        xml, text = resolve_openlp_song(church, cmap, title, sid)

        # Error categorization (string-prefix-based)
        if isinstance(text, str):
            if text.startswith("No valid song id"):
                md = append_missing(md, text_key, title, church, "invalid song ID")
                continue
            if text.startswith("Invalid hymnal prefix"):
                md = append_missing(md, text_key, title, church, "invalid hymnal prefix")
                continue
            if text.startswith("Hymn not found"):
                md = append_missing(md, text_key, title, church, "not in hymnal index")
                continue
            if text.startswith("Lyrics not found"):
                md = append_missing(md, text_key, title, church, "missing lyrics XML")
                continue

        # Normal successful case
        if text:
            md = append_below_placeholder(md, text_key, text)
        if xml:
            md = append_below_placeholder(md, xml_key, xml)

    # ----------------------
    # Closing songs (OpenLP)
    # ----------------------
    for church, cmap in (("elkton", elk_index), ("lb", lb_index)):
        title_key = f"song_closing_title_{church}"
        id_key = f"song_closing_id_{church}"
        text_key = f"song_closing_text_{church}"
        xml_key = f"song_closing_xml_{church}"

        title = extract_block(md, title_key)
        if not title:
            continue

        # LB special case: Congregational choice → we skip DB lookup.
        if church == "lb" and title.lower().startswith("congregational choice"):
            continue

        sid = extract_block(md, id_key)
        xml, text = resolve_openlp_song(church, cmap, title, sid)

        if isinstance(text, str):
            if text.startswith("No valid song id"):
                md = append_missing(md, text_key, title, church, "invalid song ID")
                continue
            if text.startswith("Invalid hymnal prefix"):
                md = append_missing(md, text_key, title, church, "invalid hymnal prefix")
                continue
            if text.startswith("Hymn not found"):
                md = append_missing(md, text_key, title, church, "not in hymnal index")
                continue
            if text.startswith("Lyrics not found"):
                md = append_missing(md, text_key, title, church, "missing lyrics XML")
                continue

        # Normal successful case
        if text:
            md = append_below_placeholder(md, text_key, text)
        if xml:
            md = append_below_placeholder(md, xml_key, xml)

    return md


# =====================================================================
# Service info block handling
# =====================================================================

def extract_serviceinfo_block(md: str) -> tuple[str | None, dict]:
    """
    Extract the serviceinfo fenced code block and parse its YAML content.
    Returns (raw_block, parsed_dict). If not found, returns (None, {}).
    """
    match = re.search(r'```serviceinfo\s*([\s\S]+?)```', md)
    if not match:
        return None, {}
    raw = match.group(1)
    try:
        data = yaml.safe_load(raw)
        if not isinstance(data, dict):
            data = {}
    except Exception:
        data = {}
    return match.group(0), data


def update_serviceinfo_block(md: str, updates: dict) -> str:
    """
    Update the serviceinfo block in md with the given updates dict.
    Returns the new markdown text.
    """
    old_block, data = extract_serviceinfo_block(md)
    data.update(updates)
    new_yaml = yaml.dump(data, sort_keys=False, allow_unicode=True)
    new_block = f"```serviceinfo\n{new_yaml}```"
    if old_block:
        return md.replace(old_block, new_block)
    else:
        # Append at end if not found
        return md.rstrip() + "\n\n" + new_block + "\n"


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

    # Update serviceinfo block with hymns and special_music
    hymns = []
    for church in ("elkton", "lb"):
        title = extract_block(updated, f"song_opening_title_{church}")
        if title:
            hymns.append(str(title).strip())
    # Optionally, add special_music from a placeholder (customize as needed)
    special_music = []
    for church in ("elkton", "lb"):
        sm_title = extract_block(updated, f"song_middle_title_{church}")
        if sm_title:
            special_music.append(str(sm_title).strip())
    updates = {}
    if hymns:
        updates["hymns"] = hymns
    if special_music:
        updates["special_music"] = special_music
    if updates:
        updated = update_serviceinfo_block(updated, updates)

    if updated != md:
        write_text(master_path, updated)
        logging.info("Updated: %s", master_path)
    else:
        logging.info("No changes made.")


if __name__ == "__main__":
    main()
