# openlp_env.py
# v1.0 – verified 2025-11-04
# Provides access to OpenLP databases, exports slide/song indexes,
# and retrieves scripture text for Elkton and LB environments.

from pathlib import Path
import sqlite3
import re

# --- Configuration -----------------------------------------------------------

# Define OpenLP root directories for SmoothJazz
OPENLP_DIRS = {
    "elkton": Path(r"C:\Users\pasto\OneDrive - United Parish of Elkton (1)\Shared with Everyone\OpenLP"),
    "lb": Path(r"C:\Users\pasto\OneDrive - United Parish of Elkton (1)\Shared with Everyone\LBopenLP")
}

# Default Bible database filenames per church
DEFAULT_BIBLES = {
    "elkton": "Good News Translation.sqlite",
    "lb": "New Revised Standard Version.sqlite"
}

# Default output directory for Markdown indexes
INDEX_DIR = Path(r"C:\Users\Public\Documents\Area42\Worship")

# --- Core Environment Helpers ------------------------------------------------

def get_root(church: str) -> Path:
    """Return the base directory for the given church."""
    church = church.lower()
    if church not in OPENLP_DIRS:
        raise ValueError(f"Unknown church '{church}'. Expected one of: {', '.join(OPENLP_DIRS)}")
    return OPENLP_DIRS[church]


def get_db_path(church: str, db_name: str) -> Path:
    """Return full path to a given OpenLP SQLite database (e.g., 'custom.sqlite')."""
    root = get_root(church)
    return root / db_name / f"{db_name}.sqlite"


def connect_db(church: str, db_name: str):
    """Open a connection to a given OpenLP database."""
    db_path = get_db_path(church, db_name)
    if not db_path.exists():
        raise FileNotFoundError(f"Database not found: {db_path}")
    conn = sqlite3.connect(db_path)
    print(f"✅ Connected to {db_path}")
    return conn

# --- List & Export Functions -------------------------------------------------

def list_custom_slides(church: str):
    """Return a list of (id, title) tuples for all custom slides."""
    conn = connect_db(church, "custom")
    cur = conn.cursor()
    cur.execute("SELECT id, title FROM custom_slide ORDER BY title;")
    results = cur.fetchall()
    conn.close()
    return results


def list_songs(church: str):
    """
    Return a list of (id, title, hymnal, entry) tuples for all songs.
    Uses songs_songbooks and song_books to include hymnals and entry numbers.
    """
    conn = connect_db(church, "songs")
    cur = conn.cursor()

    cur.execute("""
        SELECT s.id, s.title, sb.name AS hymnal, ssb.entry
        FROM songs s
        LEFT JOIN songs_songbooks ssb ON s.id = ssb.song_id
        LEFT JOIN song_books sb ON ssb.songbook_id = sb.id
        ORDER BY s.title, sb.name, ssb.entry;
    """)
    results = cur.fetchall()
    conn.close()
    return results  # (id, title, hymnal, entry)


def export_custom_slides_md(church: str, output_dir: Path | None = None):
    """Export custom slides to a Markdown index."""
    slides = list_custom_slides(church)
    output_dir = output_dir or INDEX_DIR
    md_path = output_dir / f"openlp_custom_index_{church}.md"

    with open(md_path, "w", encoding="utf-8", newline="") as f:
        f.write(f"# OpenLP Custom Slide Index — {church.title()}\n\n")
        f.write("| UUID | Title |\n|------|--------|\n")
        for sid, title in slides:
            f.write(f"| {sid} | {title} |\n")

    print(f"✅ Custom slide index written: {md_path}")
    return md_path


def export_songs_md(church: str, output_dir: Path | None = None):
    """Export songs to a Markdown file including hymnal info."""
    songs = list_songs(church)
    output_dir = output_dir or INDEX_DIR
    md_path = output_dir / f"openlp_songs_index_{church}.md"

    with open(md_path, "w", encoding="utf-8", newline="") as f:
        f.write(f"# OpenLP Songs Index — {church.title()}\n\n")
        f.write("| UUID | Title | Hymnal | Entry |\n")
        f.write("|------|--------|---------|--------|\n")
        for sid, title, hymnal, entry in songs:
            hymnal = hymnal or ""
            entry = entry or ""
            f.write(f"| {sid} | {title} | {hymnal} | {entry} |\n")

    print(f"✅ Song index (with hymnals) written: {md_path}")
    return md_path


# --- Bible Functions ---------------------------------------------------------

def get_bible_path(church: str, version: str | None = None) -> Path:
    """Return the full path to the Bible SQLite database for a given church."""
    root = get_root(church)
    bible_file = version or DEFAULT_BIBLES.get(church)
    return root / "bibles" / bible_file


def get_scripture_text(church: str, passage_ref: str) -> str:
    """
    Fetch scripture text from the church's preferred Bible translation.
    Supports multi-range references like 'John 1:1-3, 5-7; 2:8-9'.
    Returns plain text with single newlines between verses.
    """

    bible_path = get_bible_path(church)
    if not bible_path.exists():
        raise FileNotFoundError(f"Bible database not found: {bible_path}")

    passage_ref = passage_ref.replace("–", "-").replace("—", "-").strip()
    m = re.match(r"^([1-3]?\s?[A-Za-z ]+)\s+(.+)$", passage_ref)
    if not m:
        raise ValueError(f"Unable to parse passage reference: {passage_ref}")
    book = m.group(1).strip()
    remainder = m.group(2).strip()

    conn = sqlite3.connect(bible_path)
    cur = conn.cursor()

    # --- Lookup book_id ---
    cur.execute("SELECT id FROM book WHERE name LIKE ?", (book,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise ValueError(f"Book '{book}' not found in {bible_path}")
    book_id = row[0]

    verses_out = []
    last_chapter = None

    # Split by semicolon to handle chapter transitions
    for segment in remainder.split(";"):
        segment = segment.strip()
        if not segment:
            continue

        if ":" in segment:
            chapter_part, verse_part = segment.split(":", 1)
            try:
                chapter = int(chapter_part)
            except ValueError:
                chapter = last_chapter
            last_chapter = chapter
        else:
            chapter = last_chapter
            verse_part = segment

        for sub in verse_part.split(","):
            sub = sub.strip()
            if not sub:
                continue
            if "-" in sub:
                start, end = sub.split("-", 1)
                v1, v2 = int(start), int(end)
                cur.execute("""
                    SELECT text FROM verse
                    WHERE book_id=? AND chapter=? AND verse BETWEEN ? AND ?
                    ORDER BY verse
                """, (book_id, chapter, v1, v2))
            else:
                v = int(sub)
                cur.execute("""
                    SELECT text FROM verse
                    WHERE book_id=? AND chapter=? AND verse=?
                """, (book_id, chapter, v))
            verses = [row[0].strip() for row in cur.fetchall()]
            verses_out.extend(verses)

    conn.close()
    return "\n".join(verses_out)

