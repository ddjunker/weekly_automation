from pathlib import Path

def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")

def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")

def safe_mkdir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)

def list_files(path: Path, pattern: str = "*.md"):
    return sorted(path.glob(pattern))
