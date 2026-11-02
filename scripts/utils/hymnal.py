"""
Hymnal references and verse selection.

This is church naming policy rather than database access, so it lives here instead
of in openlp.py — and it is shared, so that publish and music_gather can no longer
disagree about what a song id means. They did: music_gather accepted 'r 123' and
'B218 vv 1-3,5', while publish required '^[brk]\\d{1,4}$', so a master written in
the documented form gathered lyrics correctly and then left the OpenLP service item
still called "song_opening".
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from scripts.utils.placeholder import extract_block
from scripts.utils.text_clean import clean_text

# b/r/k, optional space, 1-4 digits, and a tolerated legacy "vv 1-3,5" suffix.
SONG_ID_RE = re.compile(
    r"^(?P<prefix>[rbk])\s*(?P<entry>\d{1,4})"
    r"(?:\s*v{1,2}\.?\s*(?P<legacy_verses>[\d\s,\-]+))?$",
    re.IGNORECASE,
)

# Values that mean "this slot is deliberately not a hymn".
SKIP_MARKERS = frozenset({"tbd", "omit", "delete", "na", "congregational choice"})

HYMNALS: dict[str, dict[str, str]] = {
    "elkton": {
        "r": "United Methodist Hymnal",
        "b": "The Singing Church",
    },
    "lb": {
        "r": "United Methodist Hymnal",
        "b": "Hymns of Faith",
        "k": "The Faith We Sing",
    },
}


@dataclass(frozen=True)
class SongRef:
    """A hymnal reference: which book, which number."""

    prefix: str   # 'r' | 'b' | 'k'
    entry: str    # digits, unpadded
    raw: str      # the master's text, for messages
    legacy_verses: tuple[int, ...] | None = None  # a 'vv …' suffix, ignored but reported

    def padded(self, width: int) -> str:
        return self.entry.zfill(width)


def _strip_comment(raw: str | None) -> str:
    """Drop an Obsidian '%% … %%' note and surrounding whitespace."""
    if raw is None:
        return ""
    return clean_text(str(raw).split("%%", 1)[0]).strip()


def is_skip_marker(raw: str | None) -> bool:
    """True for 'tbd', 'omit', 'congregational choice' and friends."""
    return _strip_comment(raw).lower() in SKIP_MARKERS


def parse_verse_list(raw: str | None) -> tuple[int, ...] | None:
    """
    Parse '1-3, 5' into (1, 2, 3, 5). Returns None when no selection is given,
    which means every verse.

    Ranges are expanded: the old parser collected bare digits with a regex, so
    '1-3,5' silently became verses 1, 3 and 5 — verse 2 went missing.
    """
    text = _strip_comment(raw)
    if not text or text.lower() in SKIP_MARKERS or text.lower() == "all":
        return None

    verses: list[int] = []
    for chunk in re.split(r"[,\s]+", text):
        if not chunk:
            continue
        m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", chunk)
        if m:
            start, end = int(m.group(1)), int(m.group(2))
            verses.extend(range(min(start, end), max(start, end) + 1))
            continue
        if chunk.isdigit():
            verses.append(int(chunk))

    seen: dict[int, None] = {}
    for verse in verses:
        seen.setdefault(verse, None)
    return tuple(seen) or None


def parse_song_ref(raw: str | None, *, title_for_log: str = "") -> SongRef | None:
    """
    Parse a hymnal id: 'r 123', 'R395', 'B218 vv 1-3,5', 'k 42 vv 1, 3, 5-7'.

    Returns None when the value is not a hymnal reference at all — a performer
    name, a video, a skip marker or an empty block. Callers tell those apart with
    is_skip_marker().
    """
    cleaned = _strip_comment(raw)
    if not cleaned or is_skip_marker(cleaned):
        return None

    m = SONG_ID_RE.match(cleaned)
    if not m:
        return None

    return SongRef(
        prefix=m.group("prefix").lower(),
        entry=m.group("entry"),
        raw=cleaned,
        legacy_verses=parse_verse_list(m.group("legacy_verses")),
    )


def hymnal_name(church: str, prefix: str) -> str | None:
    """OpenLP song book name for a church's r/b/k prefix."""
    return HYMNALS.get((church or "").strip().lower(), {}).get((prefix or "").strip().lower())


def intro_slide_prefix(church: str, ref: SongRef | None, title: str) -> str | None:
    """
    Title prefix of the custom slide that introduces this song:
      - not a hymn: 'P901 <title>'
      - k: F####   r: U###   b: H### (lb) / S### (elkton)
    """
    if ref is None:
        clean_title = (title or "").strip()
        return f"P901 {clean_title}" if clean_title else None

    if ref.prefix == "k":
        return f"F{ref.padded(4)}"
    if ref.prefix == "r":
        return f"U{ref.padded(3)}"
    if ref.prefix == "b":
        head = "H" if (church or "").strip().lower() == "lb" else "S"
        return f"{head}{ref.padded(3)}"
    return None


def read_verse_filter(master_md: str, slot: str, church: str) -> tuple[int, ...] | None:
    """
    The verses to use for a slot, from {song_<slot>_vv_<church>} in the master.

    This field is the single source of verse selection: both the printed script and
    the projected song read it, so the page and the screen cannot disagree. None
    means every verse.
    """
    return parse_verse_list(extract_block(master_md, f"song_{slot}_vv_{church}"))
