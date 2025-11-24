from __future__ import annotations
"""
music_gather.py – Weekly Automation music retrieval module (Config + Utils)
---------------------------------------------------------------------------

Refactored to use:
  • scripts.utils.config.config
  • scripts.utils.placeholder (append_below_placeholder, extract_block)
  • scripts.utils.text_clean.clean_markdown
  • scripts.utils.openlp_helpers (load_openlp_song, map_song_index)

No hardcoded paths; all paths come from weekly_config.ini via config.py.
"""

import argparse
import html
import logging
import re
import sqlite3
import unicodedata
from pathlib import Path

# ---------------------------------------------------------
# Project utilities
# ---------------------------------------------------------
from scripts.utils.config import config
from scripts.utils import openlp_env  # kept for potential future use
from scripts.utils.placeholder import append_below_placeholder, extract_block
from scripts.utils.text_clean import clean_markdown
from scripts.utils.openlp_helpers import load_openlp_song, map_song_index  # :contentReference[oaicite:1]{index=1}


# =====================================================================
# Basic file helpers
# =====================================================================

def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")
    
def append_missing(md: str, placeholder: str, title: str, church: str, reason: str) -> str:
    """
    Standardized diagnostic insertion for missing songs.
    """
    msg = f"[Missing song for {church}: {title} — {reason}]"
    return append_below_placeholder(md, placeholder, msg)



# =====================================================================
# Index parsing (from markdown index files)
# =====================================================================

def parse_index(path: Path) -> list[dict]:
    """
    Parse a markdown table like openlp_songs_index_*.md into a list of dicts
    compatible with openlp_helpers.map_song_index().
    """
    rows = []
    lines = read_text(path).splitlines()
    in_table = False

    for line in lines:
        stripped = line.strip()

        if stripped.startswith("| UUID"):
            in_table = True
            continue
        if not in_table:
            continue
        if not stripped.startswith("|"):
            break

        # separator row (---- etc.)
        if set(stripped) <= {"|", "-", " "}:
            continue

        parts = [p.strip() for p in stripped.strip("|").split("|")]
        if len(parts) != 4:
            continue

        uuid_s, title, hymnal, entry_s = parts
        try:
            uuid = int(uuid_s)
        except ValueError:
            continue

        entry = entry_s.strip() or None

        rows.append({
            "uuid": uuid,
            "title": title,
            "hymnal": hymnal,
            "entry": entry,
        })

    return rows


# =====================================================================
# Hymnal mapping
# =====================================================================

def hymnal_from_prefix(church: str, prefix: str) -> str | None:
    prefix = prefix.lower()

    if church == "elkton":
        match prefix:
            case "r": return "United Methodist Hymnal"
            case "b": return "The Singing Church"
            case "k": return "The Faith We Sing"

    elif church == "lb":
        match prefix:
            case "r": return "United Methodist Hymnal"
            case "b": return "Hymns of Faith"
            case "k": return "The Faith We Sing"

    logging.error("Unknown hymnal prefix %r for %s", prefix, church)
    return None


# =====================================================================
# Song ID parsing
# =====================================================================

class ParsedSongID:
    def __init__(self, prefix: str, verses: list[int] | None):
        self.prefix = prefix
        self.verses = verses


def parse_song_id(raw: str | None) -> ParsedSongID | None:
    if raw is None:
        return None

    cleaned = raw.split("%%", 1)[0]
    cleaned = unicodedata.normalize("NFKC", cleaned)
    cleaned = (
        cleaned.replace("\u00A0", " ")
               .replace("\u200B", "")
               .replace("\u2009", " ")
               .replace("\u202F", " ")
               .strip()
    )
    if not cleaned:
        return None

    pattern = re.compile(
        r"([rbk])\s*(\d+)?(?:\s+v+v?\s+(.+))?$",
        re.IGNORECASE
    )
    m = pattern.match(cleaned)
    if not m:
        logging.warning("Could not parse song id: %r (cleaned=%r)", raw, cleaned)
        return None

    prefix = m.group(1).lower()
    verses_raw = m.group(3)

    verses = None
    if verses_raw:
        nums = re.findall(r"\d+", verses_raw)
        if nums:
            verses = [int(n) for n in nums]

    return ParsedSongID(prefix, verses)


# =====================================================================
# XML → text
# =====================================================================

def xml_to_text(xml: str) -> str:
    """
    Strip OpenLP song XML down to readable text.
    """
    text = re.sub(r"<[^>]+>", "", xml)
    text = html.unescape(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# =====================================================================
# CCLI plaintext lyrics (from config.ccli_dir)
# =====================================================================

def get_ccli_lyrics(title: str) -> str | None:
    """
    Fetch CCLI lyrics from a folder structure like:
        {ccli_dir}/{title}/{title} lyrics.txt
    """
    title = clean_markdown(title).strip()
    if not title:
        return None

    folder = config.ccli_dir / title
    file_path = folder / f"{title} lyrics.txt"
    if not file_path.exists():
        return None

    return read_text(file_path).strip()


# =====================================================================
# Text normalization for title matching
# =====================================================================

def _norm_title(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = s.replace("\u00A0", " ").replace("\u200B", "")
    s = s.strip().lower()
    return re.sub(r"\s+", " ", s)


# =====================================================================
# Core OpenLP resolution (using openlp_helpers + config)
# =====================================================================

def resolve_openlp_song(church: str, imap, title: str, id_raw: str | None):
    """
    Given a church ('elkton' or 'lb'), an index map (hymnal+entry → row),
    the song title from the Master, and the raw ID (like 'r 123 vv 1-3'),
    resolve to:
      • xml (lyrics XML string)
      • text (plain lyrics text for print)
    """
    parsed = parse_song_id(id_raw)
    if not parsed:
        return None, f"No valid song id for title {title!r}"

    hymnal = hymnal_from_prefix(church, parsed.prefix)
    if not hymnal:
        return None, f"Invalid hymnal prefix for {title}"

    search_title = _norm_title(title)

    candidates = [
        r for r in imap.values()
        if (r["hymnal"] or "").strip() == hymnal
        and _norm_title(r["title"]) == search_title
    ]

    if not candidates:
        return None, f"Hymn not found in {hymnal}: {title}"

    row = candidates[0]
    uuid = int(row["uuid"])

    # Choose DB path from config for the correct church
    db_path = config.elkton_songs_db if church == "elkton" else config.lb_songs_db

    # Use openlp_helpers to load the song row
    song_info = load_openlp_song(db_path, uuid)
    if not song_info:
        return None, f"Lyrics not found for UUID {uuid}"

    lyrics_xml = song_info["lyrics"]
    text = xml_to_text(lyrics_xml)

    if parsed.verses:
        text = "Verses: " + ", ".join(map(str, parsed.verses)) + "\n\n" + text

    return lyrics_xml.strip(), text.strip()


# =====================================================================
# Processing the Master file
# =====================================================================

def process_master(md: str, elk_map, lb_map) -> str:
    """
    Fill in song text/XML in the Master markdown using:
      • CCLI plaintext for opening songs
      • OpenLP DB + index for middle and closing songs
    """

    # Opening songs (CCLI)
    for church in ("elkton", "lb"):
        title = extract_block(md, f"song_opening_title_{church}")
        if title:
            lyr = get_ccli_lyrics(title)
            if lyr:
                md = append_below_placeholder(md, f"song_opening_text_{church}", lyr)

    # Middle songs (OpenLP)
    for church, cmap in (("elkton", elk_map), ("lb", lb_map)):
        title = extract_block(md, f"song_middle_title_{church}")
        sid = extract_block(md, f"song_middle_id_{church}")

        if title:
            xml, text = resolve_openlp_song(church, cmap, title, sid)

            if isinstance(text, str) and text.startswith("No valid song id"):
                md = append_missing(md, f"song_middle_text_{church}", title, church, "invalid song ID")
                continue

            if isinstance(text, str) and text.startswith("Invalid hymnal prefix"):
                md = append_missing(md, f"song_middle_text_{church}", title, church, "invalid hymnal prefix")
                continue

            if isinstance(text, str) and text.startswith("Hymn not found"):
                md = append_missing(md, f"song_middle_text_{church}", title, church, "not in hymnal index")
                continue

            if isinstance(text, str) and text.startswith("Lyrics not found"):
                md = append_missing(md, f"song_middle_text_{church}", title, church, "missing lyrics XML")
                continue

        # Normal successful case
            if text:
                md = append_below_placeholder(md, f"song_middle_text_{church}", text)
            if xml:
                md = append_below_placeholder(md, f"song_middle_xml_{church}", xml)


    # Closing songs
    for church, cmap in (("elkton", elk_map), ("lb", lb_map)):
        title_key = f"song_closing_title_{church}"
        id_key    = f"song_closing_id_{church}"
        text_key  = f"song_closing_text_{church}"
        xml_key   = f"song_closing_xml_{church}"

        title = extract_block(md, title_key)
        if not title:
            continue

        if church == "lb" and title.lower().startswith("congregational choice"):
            continue

        sid = extract_block(md, id_key)
        xml, text = resolve_openlp_song(church, cmap, title, sid)

        if isinstance(text, str) and text.startswith("No valid song id"):
            md = append_missing(md, f"song_closing_text_{church}", title, church, "invalid song ID")
            continue

        if isinstance(text, str) and text.startswith("Invalid hymnal prefix"):
            md = append_missing(md, f"song_closing_text_{church}", title, church, "invalid hymnal prefix")
            continue

        if isinstance(text, str) and text.startswith("Hymn not found"):
            md = append_missing(md, f"song_closing_text_{church}", title, church, "not in hymnal index")
            continue

        if isinstance(text, str) and text.startswith("Lyrics not found"):
            md = append_missing(md, f"song_closing_text_{church}", title, church, "missing lyrics XML")
            continue

        # Normal successful case
        if text:
            md = append_below_placeholder(md, f"song_closing_text_{church}", text)
        if xml:
            md = append_below_placeholder(md, f"song_closing_xml_{church}", xml)


    return md


# =====================================================================
# Main
# =====================================================================

def resolve_master_path(arg: str) -> Path:
    """
    Absolute path → use as-is.
    Relative path → config.worship_dir / arg.
    """
    p = Path(arg)
    if p.is_absolute():
        return p

    worship = config.worship_dir.expanduser().resolve()
    if not worship.exists():
        raise FileNotFoundError(f"Worship directory missing: {worship}")

    return worship / arg


def main():
    parser = argparse.ArgumentParser(description="Weekly Automation: music retrieval")
    parser.add_argument("--master", required=True)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s"
    )

    master_path = resolve_master_path(args.master)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    # Load markdown-based indexes from worship_dir
    elk_index = config.worship_dir / "openlp_songs_index_elkton.md"
    lb_index  = config.worship_dir / "openlp_songs_index_lb.md"

    elk_rows = parse_index(elk_index)
    lb_rows  = parse_index(lb_index)

    # Use openlp_helpers.map_song_index for consistent key normalization
    elk_map = map_song_index(elk_rows)
    lb_map  = map_song_index(lb_rows)

    md = clean_markdown(read_text(master_path))
    updated = process_master(md, elk_map, lb_map)

    if updated != md:
        write_text(master_path, updated)
        logging.info(f"Updated: {master_path}")
    else:
        logging.info("No changes made.")


if __name__ == "__main__":
    main()
