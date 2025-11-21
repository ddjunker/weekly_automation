import regex as re
from .text_clean import clean_text

def extract_block(md: str, placeholder: str) -> str:
    placeholder = clean_text(placeholder)
    pattern = rf"\{{{placeholder}\}}(.*?)(?=\n\{{|\Z)"
    m = re.search(pattern, md, flags=re.S)
    return m.group(1).strip() if m else ""

def has_placeholder(md: str, placeholder: str) -> bool:
    return f"{{{placeholder}}}" in md

def append_below_placeholder(md: str, placeholder: str, content: str) -> str:
    placeholder = clean_text(placeholder)
    content = clean_text(content)
    pattern = rf"(\{{{placeholder}\}}\s*)"
    return re.sub(pattern, r"\1" + content + "\n", md, count=1)

def replace_placeholder(md: str, placeholder: str, content: str) -> str:
    placeholder = clean_text(placeholder)
    content = clean_text(content)
    pattern = rf"(\{{{placeholder}\}})"
    return re.sub(pattern, content, md, count=1)
