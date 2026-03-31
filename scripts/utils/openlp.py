"""
openlp.py – Unified OpenLP database access layer
-----------------------------------------------

This module merges the functionality of the legacy `openlp_env.py`
and `openlp_helpers.py` modules into one authoritative interface.

Features:
    • All paths loaded from weekly_config.ini through utils.config
    • Consistent SQLite access for songs, custom slides, and bibles
    • Full scripture passage parser (multi-segment, multi-range)
    • Structured song index building (hymnal + entry → UUID)
    • All retrieved text cleaned using utils.text_clean
"""

from pathlib import Path
import sqlite3
import re

from scripts.utils.config import config
from scripts.utils.text_clean import clean_text, full_scrub


def sanitize_scripture_reference_for_openlp(reference: str) -> str:
        """
        Normalize scripture refs for OpenLP DB parsing.

        Strips verse-part letters immediately following numerals, e.g.
            - Genesis 12:1-4a -> Genesis 12:1-4
            - John 3:16b,17 -> John 3:16,17
        """
        cleaned = reference.replace("–", "-").replace("—", "-").strip()
        return re.sub(r"(?<=\d)[A-Za-z]+(?=(?:\s|$|[-,;]))", "", cleaned)


# ============================================================================
#  PATH HELPERS
# ============================================================================

def get_openlp_root(church: str) -> Path:
    church = church.lower()
    if church == "elkton":
        return config.elkton_root
    elif church == "lb":
        return config.lb_root
    else:
        raise ValueError(f"Unknown church '{church}'")


def get_songs_db(church: str) -> Path:
    church = church.lower()
    if church == "elkton":
        # Use a path relative to the elkton OpenLP root. Fail if missing.
        root = get_openlp_root("elkton")
        candidate = root / "songs" / "songs.sqlite"
        if not root or not candidate.exists():
            raise FileNotFoundError(f"Songs DB not found under elkton root: {candidate}")
        return candidate
    elif church == "lb":
        root = get_openlp_root("lb")
        candidate = root / "songs" / "songs.sqlite"
        if not root or not candidate.exists():
            raise FileNotFoundError(f"Songs DB not found under lb root: {candidate}")
        return candidate
    else:
        raise ValueError(f"Unknown church '{church}'")


def get_bible_db(church: str) -> Path:
    church = church.lower()
    if church == "elkton":
        # Discover a bible DB under elkton root's bibles/ directory.
        root = get_openlp_root("elkton")
        bibles_dir = root / "bibles"
        if not bibles_dir.exists():
            raise FileNotFoundError(f"Bibles directory not found under elkton root: {bibles_dir}")
        candidates = list(bibles_dir.glob("*.sqlite"))
        if not candidates:
            raise FileNotFoundError(f"No .sqlite bible files found under: {bibles_dir}")
        if len(candidates) == 1:
            return candidates[0]
        # Prefer likely filenames if multiple exist
        for name_hint in ("Good News", "GNT", "New Revised", "NRSV"):
            for c in candidates:
                if name_hint.lower() in c.name.lower():
                    return c
        # If ambiguous, prefer the first deterministically sorted
        candidates.sort(key=lambda p: p.name)
        return candidates[0]
    elif church == "lb":
        root = get_openlp_root("lb")
        bibles_dir = root / "bibles"
        if not bibles_dir.exists():
            raise FileNotFoundError(f"Bibles directory not found under lb root: {bibles_dir}")
        candidates = list(bibles_dir.glob("*.sqlite"))
        if not candidates:
            raise FileNotFoundError(f"No .sqlite bible files found under: {bibles_dir}")
        if len(candidates) == 1:
            return candidates[0]
        for name_hint in ("New Revised", "NRSV", "Revised"):
            for c in candidates:
                if name_hint.lower() in c.name.lower():
                    return c
        candidates.sort(key=lambda p: p.name)
        return candidates[0]
    else:
        raise ValueError(f"Unknown church '{church}'")


def get_custom_db(church: str) -> Path:
    root = get_openlp_root(church)
    return root / "custom" / "custom.sqlite"


# ============================================================================
#  SQLITE CONNECTOR
# ============================================================================

def connect_sqlite(path: Path) -> sqlite3.Connection:
    if not path.exists():
        raise FileNotFoundError(f"Database not found: {path}")
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


# ============================================================================
#  SONG ACCESS
# ============================================================================

def load_song(church: str, uuid: int) -> dict | None:
    """
    Return a dict with core song data and footer metadata used by OpenLP services:
    {id, title, lyrics, verse_order, hymnal, entry, authors, copyright, ccli_number}
    """
    path = get_songs_db(church)
    conn = connect_sqlite(path)
    cur = conn.cursor()

    cur.execute("""
        SELECT s.id,
               s.title,
               s.lyrics,
               s.verse_order,
               s.copyright,
               s.ccli_number,
               s.search_title,
               s.alternate_title,
               sb.name AS hymnal,
               ssb.entry AS entry,
               GROUP_CONCAT(
                   COALESCE(NULLIF(TRIM(a.display_name), ''), TRIM((COALESCE(a.first_name, '') || ' ' || COALESCE(a.last_name, '')))),
                   ', '
               ) AS authors
        FROM songs s
        LEFT JOIN songs_songbooks ssb ON s.id = ssb.song_id
        LEFT JOIN song_books sb ON ssb.songbook_id = sb.id
        LEFT JOIN authors_songs als ON s.id = als.song_id
        LEFT JOIN authors a ON als.author_id = a.id
        WHERE s.id = ?
        GROUP BY s.id, s.title, s.lyrics, s.verse_order, s.copyright, s.ccli_number, s.search_title, s.alternate_title, sb.name, ssb.entry
    """, (uuid,))

    row = cur.fetchone()
    conn.close()
    if not row:
        return None

    return {
        "id": row["id"],
        "title": clean_text(row["title"] or ""),
        "search_title": row["search_title"] or "",
        "alternate_title": row["alternate_title"] or "",
        "lyrics": full_scrub(row["lyrics"] or ""),
        "verse_order": (row["verse_order"] or "").strip(),
        "hymnal": row["hymnal"],
        "entry": row["entry"],
        "authors": clean_text(row["authors"] or ""),
        "copyright": clean_text(row["copyright"] or ""),
        "ccli_number": clean_text(str(row["ccli_number"] or "")),
    }


def load_song_by_hymnal_entry(church: str, hymnal: str, entry: str):
    """
    Return full song dict for hymnal/entry combination.
    """
    index = build_song_index(church)
    key = (hymnal, str(entry))
    uuid = index.get(key)
    if not uuid:
        return None
    return load_song(church, uuid)


_SONG_XML_TYPE_MAP = {
    "verse": "v", "chorus": "c", "bridge": "b",
    "prechorus": "p", "pre-chorus": "p", "ending": "e",
    "intro": "i", "other": "o",
}

_SONG_XML_LABEL_NAMES = {
    "v": "verse", "c": "chorus", "b": "bridge",
    "p": "pre-chorus", "e": "ending", "i": "intro",
}


def _parse_song_sections(lyrics_xml: str) -> dict[str, str] | None:
    """
    Parse OpenLyrics-style XML and return an ordered dict of section_key → text.
    Keys are like "v1", "c1", "b2" in XML declaration order.
    Returns None on parse failure or if no sections found.
    """
    import xml.etree.ElementTree as ET

    if not lyrics_xml:
        return None

    try:
        xml_str = re.sub(r"<\?xml[^?]*\?>\s*", "", lyrics_xml, count=1).strip()
        root = ET.fromstring(xml_str)
    except ET.ParseError:
        return None

    lyrics_el = root.find("lyrics")
    if lyrics_el is None:
        return None

    sections: dict[str, str] = {}
    for verse_el in lyrics_el.findall("verse"):
        vtype = (verse_el.get("type") or "v").strip().lower()
        vlabel = (verse_el.get("label") or "1").strip()
        vtype = _SONG_XML_TYPE_MAP.get(vtype, vtype)
        key = f"{vtype}{vlabel}"
        text = (verse_el.text or "").strip()
        if text and key not in sections:
            sections[key] = text

    return sections if sections else None


def parse_song_xml(lyrics_xml: str, verse_order: str | None) -> list[dict] | None:
    """
    Parse OpenLyrics-style song XML and return slides ordered by verse_order.

    Reads <verse type="..." label="..."> elements (and long-form types like
    "Verse"/"Chorus") and maps them to section keys like "v1", "c1", "b1".
    Applies verse_order (space-separated, e.g. "c1 v1 c1 v2") to produce
    the slide list with proper repetition and verseTag values.

    Returns list of {title, raw_slide, verseTag} or None if parsing fails
    (caller should fall back to xml_to_text + blank-line split).
    """
    sections = _parse_song_sections(lyrics_xml)
    if sections is None:
        return None

    order = (verse_order or "").strip().split() if verse_order else list(sections.keys())
    if not order:
        order = list(sections.keys())

    slides = []
    for key in order:
        text = sections.get(key.lower())
        if not text:
            continue
        verse_tag = key[0].upper() + key[1:]  # "c1" → "C1", "v1" → "V1"
        first_line = text.split("\n")[0].strip()
        slides.append({
            "title": (first_line or text)[:30],
            "raw_slide": text,
            "verseTag": verse_tag,
        })

    return slides if slides else None


def song_xml_to_labeled_markdown(
    lyrics_xml: str,
    verses_filter: list[int] | None = None,
) -> str | None:
    """
    Parse OpenLyrics-style song XML and return labeled markdown text.

    Each section gets a ### heading:
        v1  → ### verse 1
        v2  → ### verse 2
        c1  → ### chorus
        c2  → ### chorus 2
        b1  → ### bridge
        p1  → ### pre-chorus

    Sections are output in XML declaration order (each unique section once).
    If verses_filter is given (e.g. [1, 3]), only verse sections with those
    numbers are included; non-verse sections (chorus, bridge, etc.) are always
    included.

    Returns None on parse failure (caller should fall back to xml_to_text).
    """
    sections = _parse_song_sections(lyrics_xml)
    if sections is None:
        return None

    # Count how many distinct labels exist per type to decide if we need numbers
    type_counts: dict[str, int] = {}
    for key in sections:
        vtype = key[0]
        type_counts[vtype] = type_counts.get(vtype, 0) + 1

    parts: list[str] = []
    for key, text in sections.items():
        vtype = key[0]
        vlabel = key[1:]

        if vtype == "v":
            try:
                verse_num = int(vlabel)
            except ValueError:
                verse_num = None
            if verses_filter is not None and verse_num not in verses_filter:
                continue
            heading = f"### verse {vlabel}"
        else:
            label_name = _SONG_XML_LABEL_NAMES.get(vtype, vtype)
            if type_counts.get(vtype, 1) > 1:
                heading = f"### {label_name} {vlabel}"
            else:
                heading = f"### {label_name}"

        parts.append(f"{heading}\n{text}")

    return "\n\n".join(parts) if parts else None


def build_song_index(church: str) -> dict:
    """
    Return dict mapping (hymnal, entry) → uuid
    """
    path = get_songs_db(church)
    conn = connect_sqlite(path)
    cur = conn.cursor()

    cur.execute("""
        SELECT s.id AS uuid, sb.name AS hymnal, ssb.entry AS entry
        FROM songs s
        LEFT JOIN songs_songbooks ssb ON s.id = ssb.song_id
        LEFT JOIN song_books sb ON ssb.songbook_id = sb.id
        WHERE sb.name IS NOT NULL AND ssb.entry IS NOT NULL
    """)

    rows = cur.fetchall()
    conn.close()

    index = {}
    for row in rows:
        key = (row["hymnal"], str(row["entry"]))
        if key not in index:      # keep first occurrence
            index[key] = row["uuid"]
    return index


# ============================================================================
#  CUSTOM SLIDE ACCESS
# ============================================================================

def load_custom_slide(church: str, uuid: int) -> dict | None:
    """
    Return dict: {id, title, text}
    """
    path = get_custom_db(church)
    conn = connect_sqlite(path)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title, credits, text FROM custom_slide
        WHERE id = ?
    """, (uuid,))

    row = cur.fetchone()
    conn.close()
    if not row:
        return None

    return {
        "id": row["id"],
        "title": clean_text(row["title"] or ""),
        "credits": clean_text(row["credits"] or ""),
        "text": full_scrub(row["text"] or ""),
    }


def list_custom_slides(church: str) -> list[dict]:
    """
    Return a list of dicts:
        { "uuid": int, "title": str }
    from the church's custom slide database.
    """
    db_path = get_custom_db(church)
    conn = connect_sqlite(db_path)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title
        FROM custom_slide
        ORDER BY title COLLATE NOCASE
    """)
    rows = cur.fetchall()
    conn.close()

    slides: list[dict] = []
    for r in rows:
        # r is sqlite3.Row because connect_sqlite sets row_factory
        slide_id = int(r["id"])
        title = r["title"] or ""
        slides.append({
            "uuid": slide_id,              # keep as int for load_custom_slide()
            "title": clean_text(title),    # normalize once, consistently
        })

    return slides



# ============================================================================
#  SCRIPTURE ACCESS
# ============================================================================

def get_scripture_text(church: str, passage: str) -> str:
    """
    Full scripture parser supporting:
        • multiple segments (e.g. "John 1:1-3; 2:1-2")
        • ranges (e.g. "1-5")
        • comma-separated sequences (e.g. "1, 3, 5-7")
        • chapter transitions in comma lists (e.g. "Genesis 2:15-17, 3:1-7")
        • book names with numerals ("1 John", "2 Samuel")

    Returns cleaned text lines joined by '\n'.
    """

    db_path = get_bible_db(church)
    conn = connect_sqlite(db_path)
    cur = conn.cursor()

    # Normalize reference for OpenLP/DB parsing
    passage = sanitize_scripture_reference_for_openlp(passage)

    # Break out book vs rest
    m = re.match(r"^([1-3]?\s?[A-Za-z ]+)\s+(.+)$", passage)
    if not m:
        conn.close()
        raise ValueError(f"Unable to parse scripture reference: {passage}")

    book, remainder = m.group(1).strip(), m.group(2).strip()

    # Get book_id — try exact match first, then starts-with fallback (e.g. "Acts of the Apostles")
    cur.execute("SELECT id FROM book WHERE TRIM(name) LIKE TRIM(?)", (book,))
    row = cur.fetchone()
    if not row:
        cur.execute("SELECT id, name FROM book WHERE TRIM(name) LIKE ? ORDER BY name", (book + "%",))
        rows = cur.fetchall()
        if len(rows) == 1:
            row = rows[0]
    if not row:
        conn.close()
        raise ValueError(f"Book '{book}' not found in {db_path.name}.")
    book_id = row["id"]

    verses_out = []
    last_chapter = None

    for segment in remainder.split(";"):
        segment = segment.strip()
        if not segment:
            continue

        # Process comma-separated piece(s), each piece may include chapter context.
        for part in segment.split(","):
            part = part.strip()
            if not part:
                continue

            if ":" in part:
                chap_str, verse_expr = part.split(":", 1)
                try:
                    chapter = int(chap_str.strip())
                except ValueError:
                    chapter = last_chapter
                last_chapter = chapter
            else:
                chapter = last_chapter
                verse_expr = part

            if chapter is None:
                conn.close()
                raise ValueError(f"Unable to determine chapter in scripture reference: {passage}")

            verse_expr = verse_expr.strip()
            if not verse_expr:
                continue

            # Range a-b
            if "-" in verse_expr:
                v1_str, v2_str = verse_expr.split("-", 1)
                v1, v2 = int(v1_str.strip()), int(v2_str.strip())
                cur.execute("""
                    SELECT text FROM verse
                    WHERE book_id=? AND chapter=? AND verse BETWEEN ? AND ?
                    ORDER BY verse
                """, (book_id, chapter, v1, v2))
            else:
                v = int(verse_expr)
                cur.execute("""
                    SELECT text FROM verse
                    WHERE book_id=? AND chapter=? AND verse=?
                """, (book_id, chapter, v))

            verses = [full_scrub(r["text"]) for r in cur.fetchall()]
            verses_out.extend(verses)

    conn.close()
    return "\n".join(verses_out)
