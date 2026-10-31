"""
openlp.py - the one place that talks to OpenLP's databases and file formats.

    • Paths resolved from weekly_config.ini through utils.config
    • SQLite access for songs, custom slides and bibles
    • The scripture passage parser (multi-segment, multi-range, chapter crossing)
    • Song index building (hymnal + entry -> UUID) and OpenLyrics parsing
    • Custom slide library search, and the title/credits/footer convention OpenLP
      uses to re-link a service item to its library slide
    • All retrieved text cleaned through utils.text_clean

Service file (.osz) structure lives in utils/osz.py; hymnal naming in utils/hymnal.py.
"""

from pathlib import Path
from typing import NamedTuple
import logging
import sqlite3
import re

from scripts.utils.config import config
from scripts.utils.file_io import backup_file
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

# Preferred Bible filenames per church, most specific first. The directory is
# globbed for *.sqlite; these only break a tie when several are installed.
_BIBLE_NAME_HINTS = {
    "elkton": ("Good News", "GNT", "New Revised", "NRSV"),
    "lb": ("New Revised", "NRSV", "Revised"),
}


def get_openlp_root(church: str) -> Path:
    church = (church or "").strip().lower()
    roots = {"elkton": config.elkton_root, "lb": config.lb_root}
    if church not in roots:
        raise ValueError(f"Unknown church '{church}'")
    return roots[church]


def get_songs_db(church: str) -> Path:
    root = get_openlp_root(church)
    candidate = root / "songs" / "songs.sqlite"
    if not candidate.exists():
        raise FileNotFoundError(f"Songs DB not found under {church} root: {candidate}")
    return candidate


def get_bible_db(church: str) -> Path:
    """The church's Bible database, discovered under <root>/bibles."""
    root = get_openlp_root(church)
    bibles_dir = root / "bibles"
    if not bibles_dir.exists():
        raise FileNotFoundError(f"Bibles directory not found under {church} root: {bibles_dir}")

    candidates = sorted(bibles_dir.glob("*.sqlite"), key=lambda p: p.name)
    if not candidates:
        raise FileNotFoundError(f"No .sqlite bible files found under: {bibles_dir}")
    if len(candidates) == 1:
        return candidates[0]

    for hint in _BIBLE_NAME_HINTS.get((church or "").strip().lower(), ()):
        for candidate in candidates:
            if hint.lower() in candidate.name.lower():
                return candidate
    return candidates[0]


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
        "raw_title": row["title"] or "",
        "search_title": row["search_title"] or "",
        "alternate_title": row["alternate_title"] or "",
        "lyrics": full_scrub(row["lyrics"] or ""),
        "verse_order": (row["verse_order"] or "").strip(),
        "hymnal": row["hymnal"],
        "entry": row["entry"],
        "authors": clean_text(row["authors"] or ""),
        "raw_authors": row["authors"] or "",
        "copyright": clean_text(row["copyright"] or ""),
        "raw_copyright": row["copyright"] or "",
        "ccli_number": clean_text(str(row["ccli_number"] or "")),
        "raw_ccli_number": str(row["ccli_number"] or ""),
    }


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


def parse_song_xml(
    lyrics_xml: str,
    verse_order: str | None,
    verses_filter: tuple[int, ...] | list[int] | None = None,
) -> list[dict] | None:
    """
    Parse OpenLyrics-style song XML and return slides ordered by verse_order.

    Reads <verse type="..." label="..."> elements (and long-form types like
    "Verse"/"Chorus") and maps them to section keys like "v1", "c1", "b1".
    Applies verse_order (space-separated, e.g. "c1 v1 c1 v2") to produce
    the slide list with proper repetition and verseTag values.

    verses_filter keeps only the listed verses, after verse_order has been applied
    and only for v* sections, so choruses, bridges and tags still project. It is
    the same rule the printed script uses, so page and screen agree.

    Returns list of {title, raw_slide, verseTag} or None if parsing fails
    (caller should fall back to xml_to_text + blank-line split).
    """
    sections = _parse_song_sections(lyrics_xml)
    if sections is None:
        return None

    order = (verse_order or "").strip().split() if verse_order else list(sections.keys())
    if not order:
        order = list(sections.keys())

    wanted = {int(v) for v in verses_filter} if verses_filter else None

    slides = []
    for key in order:
        text = sections.get(key.lower())
        if not text:
            continue
        if wanted is not None:
            m = re.fullmatch(r"v(\d+)", key.lower())
            if m and int(m.group(1)) not in wanted:
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
        "raw_title": row["title"] or "",
        "credits": clean_text(row["credits"] or ""),
        "raw_credits": row["credits"] or "",
        "text": full_scrub(row["text"] or ""),
        "raw_text": row["text"] or "",
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
            "title": clean_text(title),    # normalized for search/matching only
            "raw_title": title,            # exact DB value for .osz header
        })

    return slides



def search_title_key(value: str) -> str:
    """
    Normalisation for searching the custom slide library: cleaned, casefolded and
    without commas, because CtW titles carry scripture references and a stray
    comma ("CtW Psalm 122:1-9, 3") should not decide whether a slide is found.

    Four call sites normalised slightly differently — two stripped commas and two
    did not — so prepub's check and publish's lookup could disagree about which
    slide matched.
    """
    return clean_text(value or "").lower().replace(",", "")


def find_custom_slides(church: str, prefix: str, *, exact: bool = False) -> list[dict]:
    """
    Library slides whose title matches `prefix` under search_title_key.

    exact=True requires the normalised titles to be equal. Otherwise prefix
    matches are returned, with any exact match first.
    """
    needle = search_title_key(prefix)
    if not needle:
        return []

    slides = list_custom_slides(church)
    if exact:
        return [s for s in slides if search_title_key(s["title"]) == needle]

    matches = [s for s in slides if search_title_key(s["title"]).startswith(needle)]
    matches.sort(key=lambda s: search_title_key(s["title"]) != needle)
    return matches


def get_custom_slide_pages(church: str, prefix: str, *, exact: bool = False) -> tuple[str, list[str], str] | None:
    """(title, pages, credits) for the best matching library slide, or None."""
    for match in find_custom_slides(church, prefix, exact=exact):
        slide = load_custom_slide(church, match["uuid"])
        if not slide or not str(slide.get("raw_text") or "").strip():
            continue
        pages = extract_openlp_slide_pages(slide["raw_text"])
        if pages:
            return slide["raw_title"], pages, slide.get("raw_credits", "")
    return None


def extract_openlp_slide_pages(xml_text: str) -> list[str]:
    """
    Pages of a custom slide DB record, one per CDATA section, for verbatim use in
    .osz service files. Strips the XML/CDATA wrapper but preserves OpenLP
    formatting codes ({y}, {/y}, …) so OpenLP renders each page correctly.

    A service item holds one entry per page, so callers should keep these apart
    rather than concatenating them: a joined string becomes a single overfull
    slide with the page marker printed in the middle of it.
    """
    import html as _html

    pages = [p.strip() for p in re.findall(r'<!\[CDATA\[(.*?)\]\]>', xml_text, re.DOTALL)]
    pages = [p for p in pages if p]
    if pages:
        return pages

    # Fallback: strip XML tags only, no character normalization
    text = _html.unescape(re.sub(r'<[^>]+>', '', xml_text)).strip()
    return [text] if text else []


def extract_openlp_slide_text(xml_text: str) -> str:
    """
    All pages of a custom slide as one string, separated the way OpenLP's own
    editor does it: '\\n[===]\\n'. The custom plugin splits on exactly that, so the
    marker has to stand on its own line (editcustomslideform.py splits on
    '\\n[===]\\n'); joined without the leading newline it is just literal text.
    """
    return "\n[===]\n".join(extract_openlp_slide_pages(xml_text))


def build_custom_slide_xml(slides: list[str]) -> str:
    """
    Build the XML blob OpenLP stores in custom_slide.text, one <verse> per slide.
    Matches the output of OpenLP's CustomXML.extract_xml() so the editor and the
    library search read it back unchanged.
    """
    verses = "".join(
        '<verse type="custom" label="{label}"><![CDATA[{text}]]></verse>'.format(
            label=idx,
            # Standard CDATA-splitting trick in case a slide contains "]]>".
            text=(slide or "").replace("]]>", "]]]]><![CDATA[>"),
        )
        for idx, slide in enumerate(slides, start=1)
    )
    return (
        "<?xml version='1.0' encoding='utf-8'?>\n"
        f'<song version="1.0"><lyrics language="en">{verses}</lyrics></song>'
    )


_BACKED_UP_CUSTOM_DBS: set[Path] = set()


def sync_custom_slide(church: str, title: str, credits: str, slides: list[str]) -> int | None:
    """
    Write `slides` into the library custom slide matched by title + credits, so
    OpenLP's right-click → Edit shows the same text the service file projects.
    OpenLP re-links a custom service item to this row on load (custom plugin
    service_load() matches title and credits), and the editor always reads the
    row, never the .osz.

    Inserts the row when no match exists. Returns the row id, or None when the
    text is empty or the database is unavailable.
    """
    title = (title or "").strip()
    credits = (credits or "").strip()
    slides = [s for s in (slides or []) if (s or "").strip()]
    if not title or not slides:
        return None

    try:
        db_path = get_custom_db(church)
    except Exception as e:
        logging.warning("Custom slide DB not resolvable for %s: %s", church, e)
        return None
    if not db_path.exists():
        logging.warning("Custom slide DB not found for %s: %s", church, db_path)
        return None

    xml_text = build_custom_slide_xml(slides)

    # One backup per run, before this process first writes to the library.
    if db_path not in _BACKED_UP_CUSTOM_DBS:
        _BACKED_UP_CUSTOM_DBS.add(db_path)
        try:
            backup_file(db_path, tag="sync")
        except OSError as e:
            logging.warning("Could not back up %s before writing: %s", db_path, e)

    conn = sqlite3.connect(db_path, timeout=10)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id FROM custom_slide WHERE title = ? AND IFNULL(credits, '') = ? ORDER BY id",
            (title, credits),
        )
        rows = [int(r[0]) for r in cur.fetchall()]
        if rows:
            # OpenLP links to the first match; keep any duplicates in step with it.
            cur.executemany(
                "UPDATE custom_slide SET text = ? WHERE id = ?",
                [(xml_text, row_id) for row_id in rows],
            )
            slide_id = rows[0]
        else:
            cur.execute(
                "INSERT INTO custom_slide (title, text, credits, theme_name) VALUES (?, ?, ?, '')",
                (title, xml_text, credits),
            )
            slide_id = int(cur.lastrowid)
        conn.commit()
    except sqlite3.Error as e:
        conn.rollback()
        logging.warning("Could not sync custom slide '%s' for %s: %s", title, church, e)
        return None
    finally:
        conn.close()

    return slide_id


def openlp_footer(title: str, credits: str) -> str:
    """
    Footer string for a custom service item, in OpenLP's own format:
    "<title> <credits>". The custom plugin parses credits back out as
    footer[0][len(title) + 1:], so anything else breaks the library link that
    makes right-click → Edit work.

    With no credits the footer is left empty rather than "<title> ": both parse
    back to empty credits, and an empty footer keeps the projected slide's
    footer area clear, as these slides have always had it.
    """
    title = (title or "").strip()
    credits = (credits or "").strip()
    return f"{title} {credits}" if credits else ""


def credits_from_footer(title: str, footer: list | None) -> str:
    """Inverse of openlp_footer(): recover the credits OpenLP would parse."""
    if not isinstance(footer, list) or not footer:
        return ""
    first = str(footer[0] or "")
    title = (title or "").strip()
    if title and first.startswith(title):
        return first[len(title) + 1:].strip()
    return first.strip()


# ============================================================================
#  SCRIPTURE ACCESS
# ============================================================================

class Verse(NamedTuple):
    """One verse row: chapter, verse number, cleaned text."""

    chapter: int
    verse: int
    text: str


def _verse_number(raw: str, passage: str) -> int:
    try:
        return int(raw.strip())
    except ValueError:
        raise ValueError(f"Unable to parse verse number {raw.strip()!r} in: {passage}") from None


def fetch_scripture_verses(church: str, passage: str) -> list[Verse]:
    """
    Full scripture parser supporting:
        • multiple segments (e.g. "John 1:1-3; 2:1-2")
        • ranges (e.g. "1-5") including ones that cross a chapter ("1:1-2:4")
        • comma-separated sequences (e.g. "1, 3, 5-7")
        • chapter transitions in comma lists (e.g. "Genesis 2:15-17, 3:1-7")
        • book names with numerals ("1 John", "2 Samuel")

    This is the single implementation: publish.py used to carry its own copy that
    returned (chapter, verse, text) tuples, and the two had already drifted apart
    over how a malformed chapter number is handled.
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

    verses_out: list[Verse] = []
    last_chapter = None

    try:
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
                    raise ValueError(f"Unable to determine chapter in scripture reference: {passage}")

                verse_expr = verse_expr.strip()
                if not verse_expr:
                    continue

                # Range a-b (may cross chapters, e.g. "1-2:4" from "Genesis 1:1-2:4a")
                if "-" in verse_expr:
                    v1_str, v2_str = verse_expr.split("-", 1)
                    v1 = _verse_number(v1_str, passage)
                    v2_str = v2_str.strip()
                    if ":" in v2_str:
                        end_chap_str, end_verse_str = v2_str.split(":", 1)
                        end_chapter = _verse_number(end_chap_str, passage)
                        v2 = _verse_number(end_verse_str, passage)
                        cur.execute("""
                            SELECT chapter, verse, text FROM verse
                            WHERE book_id=?
                              AND ((chapter = ? AND verse >= ?)
                                   OR (chapter > ? AND chapter < ?)
                                   OR (chapter = ? AND verse <= ?))
                            ORDER BY chapter, verse
                        """, (book_id, chapter, v1, chapter, end_chapter, end_chapter, v2))
                        last_chapter = end_chapter
                    else:
                        v2 = _verse_number(v2_str, passage)
                        cur.execute("""
                            SELECT chapter, verse, text FROM verse
                            WHERE book_id=? AND chapter=? AND verse BETWEEN ? AND ?
                            ORDER BY verse
                        """, (book_id, chapter, v1, v2))
                else:
                    v = _verse_number(verse_expr, passage)
                    cur.execute("""
                        SELECT chapter, verse, text FROM verse
                        WHERE book_id=? AND chapter=? AND verse=?
                    """, (book_id, chapter, v))

                verses_out.extend(
                    Verse(int(r["chapter"]), int(r["verse"]), full_scrub(r["text"] or ""))
                    for r in cur.fetchall()
                )
    finally:
        conn.close()

    return verses_out


def get_scripture_text(church: str, passage: str) -> str:
    """The passage as cleaned text lines joined by newlines."""
    return "\n".join(v.text for v in fetch_scripture_verses(church, passage))
