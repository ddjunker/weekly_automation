from pathlib import Path

"""
file_io.py – Simple wrappers for filesystem operations
-------------------------------------------------------

This module provides small, convenience-level helpers for reading and writing
text files, creating directories, and listing files. It does *not* perform
any text cleaning or path resolution—that work is handled by `text_clean.py`
and `config.py` respectively.

Purpose:
    • Centralize UTF-8 file reads/writes (with errors="replace" on read).
    • Reduce duplicated boilerplate in gather scripts.
    • Provide safe directory creation (`mkdir(..., exist_ok=True)`).
    • Provide a simple file-listing helper using pathlib/glob.

What this module does NOT do:
    • It does not clean, normalize, or modify text content.
    • It does not validate file types or parse formats.
    • It does not hold path configuration (use config.py for that).
    • It does not interact with OpenLP or any external systems.

Why it exists:
    These wrappers keep the gather scripts cleaner and more readable,
    reducing repetitive low-level filesystem code. They also help ensure
    consistent UTF-8 handling across all scripts in the Weekly Automation
    project.

Functions:
    read_text(path: Path) -> str
        Reads a text file using UTF-8 and replaces invalid characters.
    write_text(path: Path, text: str) -> None
        Writes text to a file using UTF-8.
    safe_mkdir(path: Path) -> None
        Creates a directory tree if it does not already exist.
    list_files(path: Path, pattern: str="*.md")
        Returns a sorted list of matching files under the directory.

This module is intentionally lightweight and generic so it can remain stable
even as text handling or configuration evolves elsewhere in the project.
"""

def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")

def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")

def safe_mkdir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)

def list_files(path: Path, pattern: str = "*.md"):
    return sorted(path.glob(pattern))
