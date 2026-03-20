import yaml
import regex as re
from .text_clean import clean_text


MARKDOWN_COMMENT_RE = re.compile(r"<!--[\s\S]*?-->")
OBSIDIAN_COMMENT_RE = re.compile(r"%%[\s\S]*?%%")


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
    # lookahead also handles =={next_placeholder}== lines
    pattern = rf"\{{{placeholder}\}}=*\s*(.*?)(?=\n\s*=*\{{|\Z)"
    m = re.search(pattern, md, flags=re.S)
    return strip_markdown_comments(m.group(1)) if m else ""

def has_placeholder(md: str, placeholder: str) -> bool:
    return f"{{{placeholder}}}" in md

def append_below_placeholder(md: str, placeholder: str, content: str) -> str:
    placeholder = clean_text(placeholder)
    content = clean_text(content)
    # =* after token handles Obsidian highlight syntax: =={placeholder}==
    pattern = rf"(\{{{placeholder}\}}=*\s*)"
    return re.sub(pattern, r"\g<1>" + content + "\n", md, count=1)

def replace_placeholder(md: str, placeholder: str, content: str) -> str:
    placeholder = clean_text(placeholder)
    content = clean_text(content)
    pattern = rf"(\{{{placeholder}\}})"
    return re.sub(pattern, content, md, count=1)


# ---------------------------------------------------------------------------
# Serviceinfo block helpers (shared by text_gather.py and music_gather.py)
# ---------------------------------------------------------------------------

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
    # Append at end if not found
    return md.rstrip() + "\n\n" + new_block + "\n"
