import logging

try:
    import regex as re
except ImportError:
    import re  # type: ignore[no-redef]
from .text_clean import clean_text


MARKDOWN_COMMENT_RE = re.compile(r"<!--[\s\S]*?-->")
OBSIDIAN_COMMENT_RE = re.compile(r"%%[\s\S]*?%%")

# A placeholder name is at least 3 characters. Every OpenLP inline formatting tag
# is 1-2 ({h} {y} {p} {it} {st}), and the shortest real placeholder in the worship
# tree is {001} in the funeral templates — so the class must allow a leading digit.
TOKEN_NAME = r"[A-Za-z0-9_]{3,}"

# What ends a block in the master: a placeholder alone on its line, allowing the
# Obsidian highlight form =={token}== and a trailing %% comment %%. Anchoring on a
# bare "{" instead would end the block at an OpenLP tag such as {h}Red Hymnal #395{/h},
# silently truncating the song text that follows it.
# The trailing \r? matters: on a CRLF master, $ matches before the \n but after the
# \r, so without it the whole terminator fails and blocks run to the end of the file.
BLOCK_END = rf"^[ \t]*=*\{{{TOKEN_NAME}\}}=*[ \t]*(?:%%.*?%%)?[ \t]*\r?$"


def strip_markdown_comments(text: str) -> str:
    """Remove Markdown/Obsidian comments from text."""
    if not text:
        return ""
    text = MARKDOWN_COMMENT_RE.sub("", text)
    text = OBSIDIAN_COMMENT_RE.sub("", text)
    return text.strip()

def extract_block(md: str, placeholder: str) -> str:
    placeholder = clean_text(placeholder)
    # =* after token handles Obsidian highlight syntax: =={placeholder}==
    # The lookahead is line-anchored so blank content does not consume the next
    # placeholder line (works for LF and CRLF files).
    pattern = rf"\{{{re.escape(placeholder)}\}}=*[ \t]*(.*?)(?={BLOCK_END}|\Z)"
    m = re.search(pattern, md, flags=re.S | re.M)
    return strip_markdown_comments(m.group(1)) if m else ""

def has_placeholder(md: str, placeholder: str) -> bool:
    return f"{{{placeholder}}}" in md

class PlaceholderNotFound(KeyError):
    """Raised when a master has no such placeholder, so the caller can report it."""


def write_block(md: str, placeholder: str, content: str, *, mode: str = "replace") -> str:
    """
    Set a placeholder's block in the master.

    mode="replace" (the default) clears whatever is currently under the token
    first, so running a gather twice leaves one copy of the text rather than two.
    mode="append" keeps the old behaviour of adding below what is there.

    Raises PlaceholderNotFound when the token is absent — previously the text was
    silently dropped on the floor and the run still reported success.
    """
    placeholder = clean_text(placeholder)
    content = clean_text(content)

    token = re.search(rf"\{{{re.escape(placeholder)}\}}=*[ \t]*", md)
    if token is None:
        raise PlaceholderNotFound(placeholder)

    if mode == "append":
        # Insert below the token's line, never on it: text glued onto the token line
        # stops that token from ending the block above it (BLOCK_END needs the token
        # alone on its line), so {ctw_ref_elkton} swallowed the whole CtW text.
        eol = md.find("\n", token.end())
        if eol == -1:
            return md + "\n" + content + "\n"
        return md[: eol + 1] + content + "\n" + md[eol + 1:]

    end = re.search(BLOCK_END, md[token.end():], flags=re.M)
    body_end = token.end() + (end.start() if end else len(md) - token.end())
    tail = md[body_end:]
    return md[: token.end()] + "\n" + content + "\n\n" + tail.lstrip("\n")


def append_below_placeholder(md: str, placeholder: str, content: str) -> str:
    """
    Deprecated: use write_block(). Kept so existing callers keep working, including
    the diagnostic messages that genuinely want to accumulate several lines.
    """
    try:
        return write_block(md, placeholder, content, mode="append")
    except PlaceholderNotFound:
        logging.warning("Placeholder {%s} not in the master; nothing written", placeholder)
        return md

_OMIT_MARKERS = {"omit", "delete"}


def omitted_marker(value: str) -> str | None:
    """
    Return the omit/delete marker a master block carries, or None.

    The marker is recognised on the block's first non-empty line, so a block left
    as "omit" with notes under it still counts as omitted. Requiring the whole
    block to strip to exactly "omit" meant that "omit" followed by anything at all
    was published verbatim.
    """
    for line in (value or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        stripped = line.strip().lower()
        if not stripped:
            continue
        return stripped if stripped in _OMIT_MARKERS else None
    return None


def master_value(md: str, placeholder: str) -> str:
    """
    A master block's value, with 'omit'/'delete' resolved to an empty string.

    Use this anywhere a master value drives a lookup or an injection: reading the
    raw block instead made 'omit' a literal search term, which is why an omitted
    {aof_ref} sent publish and prepub hunting for a slide called "AoF pomit".
    """
    value = extract_block(md, placeholder).strip()
    return "" if omitted_marker(value) else value
