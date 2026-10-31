#!/usr/bin/env python3
"""
music_gather.py - fills the song text placeholders in the master markdown.

Each of the three slots (opening, middle, closing) is routed by its id, for both
churches:

    {song_<slot>_id_<church>} parses as a hymnal reference (r/b/k + number)
        -> OpenLP songs database, rendered as labelled verse/chorus markdown
    anything else (a performer name, a video)
        -> CCLI plain-text lyrics from config.ccli_dir, by title
    a skip marker ('tbd', 'omit', 'congregational choice')
        -> left alone

Verse selection comes from {song_<slot>_vv_<church>}; publish reads the same field
for the projected song, so the printed script and the screen agree.

Writes are idempotent: re-running replaces a slot's text rather than appending a
second copy, and the master is backed up first.
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.utils.cli import resolve_master_or_exit, start_logging
from scripts.utils.config import config
from scripts.utils.file_io import backup_file, read_text, write_text, write_text_atomic
from scripts.utils.placeholder import (
    PlaceholderNotFound,
    extract_block,
    write_block,
)
from scripts.utils.text_clean import clean_markdown, clean_text, xml_to_text
from scripts.utils.hymnal import (
    hymnal_name,
    is_skip_marker,
    parse_song_ref,
    read_verse_filter,
)
from scripts.utils.openlp import build_song_index, load_song, song_xml_to_labeled_markdown



# =====================================================================
# Standardized diagnostic insertion
# =====================================================================

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

    return label_ccli_song_parts(clean_markdown(raw))


# A CCLI section label alone on its line: "Verse 1", "Chorus", "Pre-Chorus 2",
# "BRIDGE", "Tag:". Lyric lines that merely start with one of these words
# ("Chorus of angels singing") are not labels and stay as they are.
_CCLI_PART_LABEL_RE = re.compile(
    r"^(?:verse|chorus|pre-?chorus|post-?chorus|bridge|ending|tag|intro|outro"
    r"|interlude|refrain|coda|instrumental|vamp)(?:\s+\d+)?\s*:?$",
    re.IGNORECASE,
)


def label_ccli_song_parts(text: str) -> str:
    """Turn CCLI section labels into ### headings, matching the hymn markdown."""
    lines = []
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        stripped = line.strip()
        if _CCLI_PART_LABEL_RE.match(stripped):
            line = f"### {stripped}"
        lines.append(line)
    return "\n".join(lines)


# =====================================================================
# Hymnal mapping
# =====================================================================

def hymnal_from_prefix(church: str, prefix: str) -> str | None:
    """Deprecated alias for utils.hymnal.hymnal_name, kept for existing callers."""
    name = hymnal_name(church, prefix)
    if not name:
        logging.error("Unknown hymnal prefix %r for %s", prefix, church)
    return name


# =====================================================================
# XML → plain text
# =====================================================================



def resolve_openlp_song(
    church: str,
    index: dict,
    title: str,
    id_raw: str | None,
    verses_filter: tuple[int, ...] | None = None,
):
    """
    Given:
        church: 'elkton' or 'lb'
        index:  (hymnal, entry) → uuid (from build_song_index(church))
        title:  song title from Master (for diagnostics)
        id_raw: like 'r 123'
        verses_filter: verses to keep, from {song_<slot>_vv_<church>}

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

    parsed = parse_song_ref(id_raw, title_for_log=title)
    if not parsed:
        return f"No valid song id for title {title!r}"

    hymnal = hymnal_name(church, parsed.prefix)
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

    verses = list(verses_filter) if verses_filter else None
    text = song_xml_to_labeled_markdown(lyrics_xml, verses_filter=verses)
    if text is None:
        text = xml_to_text(lyrics_xml)
        if verses:
            verse_note = "Verses: " + ", ".join(map(str, verses))
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


def _gather_song_text(church: str, index: dict, title: str, id_raw: str | None,
                      verses_filter: tuple[int, ...] | None = None) -> str:
    """
    Route a single song slot to hymnal or praise-song lookup.

    If id_raw parses as a hymnal reference (r/b/k + digits), look it up in
    the OpenLP DB and return labeled markdown.  Otherwise treat the slot as a
    praise song and fetch CCLI plain-text lyrics by title.

    Always returns a non-empty string; error strings begin with "[Missing".
    """
    if parse_song_ref(id_raw or "", title_for_log=title):
        result = resolve_openlp_song(church, index, title, id_raw, verses_filter)
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

    Verse selection comes from {song_<slot>_vv_<church>}, the same field publish
    uses for the projected song, so the printed script and the screen agree.
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

            # Skip markers (congregational choice, tbd, omit) mean "no lookup".
            if is_skip_marker(title) or is_skip_marker(sid):
                continue

            text = _gather_song_text(
                church, index, title, sid, read_verse_filter(md, slot, church)
            )
            if text:
                try:
                    md = write_block(md, text_key, text)
                except PlaceholderNotFound:
                    logging.warning("Placeholder {%s} not in the master; song text not written", text_key)

    return md



# =====================================================================
# Dry-run check
# =====================================================================

def _has_unfilled_placeholder(val: str) -> bool:
    """Return True if val contains an unresolved {placeholder} token."""
    return bool(re.search(r'\{[^}]+\}', val))


def check_master(md: str, elk_index: dict, lb_index: dict) -> list[dict]:
    """Check song availability without modifying md. Returns list of result dicts."""
    results = []
    for slot in ("opening", "middle", "closing"):
        for church, index in (("elkton", elk_index), ("lb", lb_index)):
            title = extract_block(md, f"song_{slot}_title_{church}").strip()
            sid = extract_block(md, f"song_{slot}_id_{church}").strip()

            # Treat unresolved placeholder tokens as empty
            if _has_unfilled_placeholder(title):
                title = ""
            if _has_unfilled_placeholder(sid):
                sid = ""

            if not title and not sid:
                results.append({
                    "slot": slot,
                    "church": church,
                    "title": "—",
                    "id": "—",
                    "status": "EMPTY",
                    "detail": "No title/ID provided",
                })
                continue

            if is_skip_marker(title) or is_skip_marker(sid):
                marker = title or sid
                results.append({"slot": slot, "church": church, "title": title or "—",
                                "id": sid or "—", "status": "skipped", "detail": marker})
                continue

            if parse_song_ref(sid or "", title_for_log=title):
                result_text = resolve_openlp_song(
                    church, index, title, sid, read_verse_filter(md, slot, church)
                )
                status = "found"
                detail = "Hymnal DB"
                for prefix, reason in _OPENLP_ERROR_REASONS.items():
                    if result_text.startswith(prefix):
                        status = "MISSING"
                        detail = reason
                        break
            else:
                lyr = get_ccli_lyrics(clean_text(title))
                status = "found" if lyr else "MISSING"
                detail = "CCLI" if lyr else "CCLI lyrics not found"

            results.append({"slot": slot, "church": church, "title": title,
                            "id": sid or "CCLI", "status": status, "detail": detail})
    return results


def _cell(val: str) -> str:
    """Sanitize a value for use in a markdown table cell."""
    return val.replace("\n", " ").replace("|", "\\|").strip()


def _write_music_check_report(master_path: Path, results: list[dict]) -> Path:
    """Write dry-run check results to a markdown report file next to the master."""
    report_path = master_path.with_name(master_path.stem + "_music_check.md")
    lines = [f"# Music Check: {master_path.name}\n"]
    lines.append("| Slot | Church | Title | ID | Status | Detail |")
    lines.append("|------|--------|-------|----|--------|--------|")
    for r in results:
        lines.append(
            f"| {_cell(r['slot'])} | {_cell(r['church'])} | {_cell(r['title'])} "
            f"| {_cell(r['id'])} | {_cell(r['status'])} | {_cell(r['detail'])} |"
        )
    lines.append("")
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


# =====================================================================
# Main
# =====================================================================

def main():
    parser = argparse.ArgumentParser(description="Weekly Automation: music retrieval")
    parser.add_argument(
        "master",
        nargs="?",
        default="",
        help="Master markdown filename or path",
    )
    parser.add_argument("--verbose", action="store_true", help="Enable debug logging")
    args = parser.parse_args()

    start_logging(args)

    master_path = resolve_master_or_exit(args, script="music_gather.py")

    # Build song indexes once per church from OpenLP DB
    logging.info("Building OpenLP song indexes...")
    elk_index = build_song_index("elkton")
    lb_index = build_song_index("lb")

    md = read_text(master_path)

    updated = process_master(md, elk_index, lb_index)

    if updated != md:
        backup = backup_file(master_path, tag="music_gather")
        write_text_atomic(master_path, updated)
        logging.info("Updated: %s", master_path)
        if backup:
            logging.info("Previous version: %s", backup)
    else:
        logging.info("No changes made.")


if __name__ == "__main__":
    main()
