"""
Unified text cleaning utilities.
Merged from text_clean.py and unicode_scrub.py
"""

import re
import unicodedata
from textwrap import dedent


# ---------------------------------------------------------------------------
# Invisible character maps
# ---------------------------------------------------------------------------

ZERO_WIDTH_CHARS = [
    '\u200b',  # zero-width space
    '\u200c',  # zero-width non-joiner
    '\u200d',  # zero-width joiner
    '\u2060',  # word joiner
    '\ufeff',  # byte-order mark
]

# map of smart punctuation → ASCII equivalents
SMART_MAP = {
    "“": '"',
    "”": '"',
    "‘": "'",
    "’": "'",
    "—": "-",
    "–": "-",
    "\u00a0": " ",  # no-breaking space
    "\u2009": " ",  # thin space
}


smart_re = re.compile("|".join(re.escape(k) for k in SMART_MAP.keys()))


# ---------------------------------------------------------------------------
# Basic cleaning helpers
# ---------------------------------------------------------------------------

def strip_zero_width(text: str) -> str:
    for ch in ZERO_WIDTH_CHARS:
        text = text.replace(ch, "")
    return text


def normalize_smart_punctuation(text: str) -> str:
    return smart_re.sub(lambda m: SMART_MAP[m.group(0)], text)


def normalize_unicode(text: str) -> str:
    """NFC normalization + remove oddball chars."""
    text = unicodedata.normalize("NFC", text)
    text = strip_zero_width(text)
    text = normalize_smart_punctuation(text)
    return text


def normalize_whitespace(text: str) -> str:
    """Collapse irregular whitespace but preserve double-newline structure."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Single-line whitespace normalization
    lines = [re.sub(r"[ \t]+", " ", line).rstrip() for line in text.split("\n")]
    return "\n".join(lines)


def dedent_strip(text: str) -> str:
    return dedent(text).strip()


# ---------------------------------------------------------------------------
# Markdown-specific cleaners
# ---------------------------------------------------------------------------

def collapse_blank_lines(text: str) -> str:
    """Remove excessive blank lines but preserve paragraph breaks."""
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def clean_markdown_blocks(text: str) -> str:
    """
    Clean markdown block elements (`## Heading`, lists, etc)
    without disturbing their meaning.
    """
    # Remove trailing spaces on headings
    text = re.sub(r"^(#+ .+?)\s+$", r"\1", text, flags=re.MULTILINE)

    # Fix bullet indent issues
    text = re.sub(r"^\s*[-*]\s+", lambda m: m.group(0).strip() + " ", text, flags=re.MULTILINE)

    return text


# ---------------------------------------------------------------------------
# High-level clean functions
# ---------------------------------------------------------------------------

def clean_text(text: str) -> str:
    """
    General text cleaner for ALL gather scripts.
    Removes unicode junk, normalizes spacing.
    """
    if not text:
        return ""
    text = normalize_unicode(text)
    text = normalize_whitespace(text)
    return text.strip()


def clean_markdown(text: str) -> str:
    """
    Markdown-aware cleaner.
    This does NOT reflow paragraphs (belongs to dissemination),
    but it does ensure safe characters and spacing.
    """
    if not text:
        return ""
    text = clean_text(text)
    text = collapse_blank_lines(text)
    text = clean_markdown_blocks(text)
    return text.strip()


# ---------------------------------------------------------------------------
# Optional helpers (useful for OpenLP, CCLI, BibleGateway quirks)
# ---------------------------------------------------------------------------

def remove_html_entities(text: str) -> str:
    """Convert common HTML entities."""
    ENTITY_MAP = {
        "&nbsp;": " ",
        "&amp;": "&",
        "&quot;": '"',
        "&apos;": "'",
    }
    for k, v in ENTITY_MAP.items():
        text = text.replace(k, v)
    return text


def full_scrub(text: str) -> str:
    """
    Highest-level scrub: unicode cleaning, smart punctuation,
    whitespace, HTML entities, and block cleanup.
    """
    if not text:
        return ""

    text = remove_html_entities(text)
    text = clean_markdown(text)
    return text
