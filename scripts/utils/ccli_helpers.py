from pathlib import Path
from .text_clean import clean_text

def get_ccli_lyrics(title: str, root: Path) -> str | None:
    title = clean_text(title)
    folder = root / title
    file_path = folder / f"{title} lyrics.txt"
    if not file_path.exists():
        return None
    return clean_text(file_path.read_text(encoding="utf-8"))
