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
    pattern = rf"\{{{placeholder}\}}\s*(.*?)(?=\n\s*\{{|\Z)"
    m = re.search(pattern, md, flags=re.S)
    return strip_markdown_comments(m.group(1)) if m else ""

def has_placeholder(md: str, placeholder: str) -> bool:
    return f"{{{placeholder}}}" in md

def append_below_placeholder(md: str, placeholder: str, content: str) -> str:
    placeholder = clean_text(placeholder)
    content = clean_text(content)
    pattern = rf"(\{{{placeholder}\}}\s*)"
    return re.sub(pattern, r"\g<1>" + content + "\n", md, count=1)

def replace_placeholder(md: str, placeholder: str, content: str) -> str:
    placeholder = clean_text(placeholder)
    content = clean_text(content)
    pattern = rf"(\{{{placeholder}\}})"
    return re.sub(pattern, content, md, count=1)
