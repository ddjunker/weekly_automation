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
    Return a dict: {id, title, lyrics, hymnal, entry}
    """
    path = get_songs_db(church)
    conn = connect_sqlite(path)
    cur = conn.cursor()

    cur.execute("""
        SELECT s.id, s.title, s.lyrics,
               sb.name AS hymnal, ssb.entry AS entry
        FROM songs s
        LEFT JOIN songs_songbooks ssb ON s.id = ssb.song_id
        LEFT JOIN song_books sb ON ssb.songbook_id = sb.id
        WHERE s.id = ?
    """, (uuid,))

    row = cur.fetchone()
    conn.close()
    if not row:
        return None

    return {
        "id": row["id"],
        "title": clean_text(row["title"] or ""),
        "lyrics": full_scrub(row["lyrics"] or ""),
        "hymnal": row["hymnal"],
        "entry": row["entry"],
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
        SELECT id, title, text FROM custom_slide
        WHERE id = ?
    """, (uuid,))

    row = cur.fetchone()
    conn.close()
    if not row:
        return None

    return {
        "id": row["id"],
        "title": clean_text(row["title"] or ""),
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
        • book names with numerals ("1 John", "2 Samuel")

    Returns cleaned text lines joined by '\n'.
    """

    db_path = get_bible_db(church)
    conn = connect_sqlite(db_path)
    cur = conn.cursor()

    # Normalize dash types
    passage = passage.replace("–", "-").replace("—", "-").strip()

    # Break out book vs rest
    m = re.match(r"^([1-3]?\s?[A-Za-z ]+)\s+(.+)$", passage)
    if not m:
        conn.close()
        raise ValueError(f"Unable to parse scripture reference: {passage}")

    book, remainder = m.group(1).strip(), m.group(2).strip()

    # Get book_id
    cur.execute("SELECT id FROM book WHERE name LIKE ?", (book,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise ValueError(f"Book '{book}' not found in DB.")
    book_id = row["id"]

    verses_out = []
    last_chapter = None

    for segment in remainder.split(";"):
        segment = segment.strip()
        if not segment:
            continue

        # Determine chapter
        if ":" in segment:
            chap_str, verse_expr = segment.split(":", 1)
            try:
                chapter = int(chap_str)
            except ValueError:
                chapter = last_chapter
            last_chapter = chapter
        else:
            chapter = last_chapter
            verse_expr = segment

        # Process comma-separated piece(s)
        for part in verse_expr.split(","):
            part = part.strip()
            if not part:
                continue

            # Range a-b
            if "-" in part:
                v1, v2 = map(int, part.split("-", 1))
                cur.execute("""
                    SELECT text FROM verse
                    WHERE book_id=? AND chapter=? AND verse BETWEEN ? AND ?
                    ORDER BY verse
                """, (book_id, chapter, v1, v2))
            else:
                v = int(part)
                cur.execute("""
                    SELECT text FROM verse
                    WHERE book_id=? AND chapter=? AND verse=?
                """, (book_id, chapter, v))

            verses = [full_scrub(r["text"]) for r in cur.fetchall()]
            verses_out.extend(verses)

    conn.close()
    return "\n".join(verses_out)
