#!/usr/bin/env python3
"""
Weekly Automation: text_gather

- Fills in scripture text blocks for Elkton and LB from the local OpenLP Bible
  databases via openlp_env.get_scripture_text().
- Optionally captures Offertory and Commission & Benediction text by opening
  a Firefox window and letting the user copy from their preferred websites.

Usage:
    python text_gather.py --master "Master 2025-11-23.md" [--no-browser]

This script assumes the "master" markdown file contains lines like:

    {scripture_ref_1}
    Psalm 123

    {scripture_text_1_elkton}
    {scripture_text_1_lb}

and single placeholders for:

    {offertory_text}
    {benediction}
"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import List, Tuple

from utils import openlp_env

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
    import pyperclip

    HAVE_BROWSER_DEPS = True
except Exception:  # ImportError or others
    sync_playwright = None  # type: ignore
    PlaywrightTimeoutError = Exception  # type: ignore
    pyperclip = None  # type: ignore
    HAVE_BROWSER_DEPS = False


# ---------------------------------------------------------------------------
# Basic helpers
# ---------------------------------------------------------------------------

# Adjust if you ever move the worship folder; for now this matches your setup.
WORSHIP_DIR = Path(r"C:\Users\Public\Documents\Area42\Worship")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


def clean_text(s: str) -> str:
    """
    Very light normalisation:
    - normalise newlines
    - strip leading/trailing whitespace
    - collapse runs of blank lines to at most two.
    """
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in s.split("\n")]
    cleaned: List[str] = []
    blank_count = 0
    for ln in lines:
        if ln.strip():
            cleaned.append(ln)
            blank_count = 0
        else:
            blank_count += 1
            if blank_count <= 2:
                cleaned.append("")
    return "\n".join(cleaned).strip()


# ---------------------------------------------------------------------------
# Placeholder helpers
# ---------------------------------------------------------------------------

def append_below_placeholder(md: str, placeholder: str, block: str) -> str:
    """
    Insert `block` (as a markdown paragraph) immediately *after* the line that
    contains `{placeholder}`. If the placeholder is not found, return md.
    """
    pattern = re.compile(rf"^.*\{{{re.escape(placeholder)}}}.*$", re.MULTILINE)
    m = pattern.search(md)
    if not m:
        logging.warning("Placeholder {%s} not found; cannot append.", placeholder)
        return md

    insert_pos = m.end()
    block = clean_text(block)
    if not block:
        return md

    insert_segment = "\n\n" + block + "\n\n"
    return md[:insert_pos] + insert_segment + md[insert_pos:]


def replace_placeholder_with_block(md: str, placeholder: str, block: str) -> str:
    """
    Replace a line that consists solely of `{placeholder}` with `block`.

    If the placeholder is not present, the original string is returned.
    """
    pattern = re.compile(rf"^\s*\{{{re.escape(placeholder)}}}\s*$", re.MULTILINE)
    block = clean_text(block)
    if not block:
        return md
    new_md, count = pattern.subn(block, md, count=1)
    if count == 0:
        logging.warning("Placeholder {%s} not found; cannot replace.", placeholder)
        return md
    return new_md


# ---------------------------------------------------------------------------
# Scripture handling (via openlp_env.get_scripture_text)
# ---------------------------------------------------------------------------

SCRIPTURE_REF_RE = re.compile(r"\{scripture_ref_(\d+)\}")


def find_scripture_requests(md: str) -> List[Tuple[int, str]]:
    """
    Scan the master markdown for patterns like:

        {scripture_ref_1}
        Psalm 123

    and return a list of (n, reference_string).
    """
    lines = md.splitlines()
    requests: List[Tuple[int, str]] = []

    for i, line in enumerate(lines):
        m = SCRIPTURE_REF_RE.search(line.strip())
        if not m:
            continue
        n = int(m.group(1))

        # Look for the first non-empty line *after* the placeholder
        ref = ""
        for j in range(i + 1, len(lines)):
            candidate = lines[j].strip()
            if candidate:
                ref = candidate
                break

        if ref:
            requests.append((n, ref))
        else:
            logging.warning("No reference text found after {scripture_ref_%d}", n)

    return requests


def insert_scripture_blocks(md: str) -> str:
    """
    For each {scripture_ref_n} / reference pair, fetch scripture for Elkton and LB
    and replace:

        {scripture_text_n_elkton}
        {scripture_text_n_lb}
    """
    requests = find_scripture_requests(md)
    if not requests:
        logging.info("Found 0 scripture references.")
        return md

    logging.info("Found %d scripture reference(s).", len(requests))

    for n, ref in requests:
        elk_ph = f"scripture_text_{n}_elkton"
        lb_ph = f"scripture_text_{n}_lb"
        logging.info("Fetching scripture %d: %s", n, ref)

        try:
            elk_text = openlp_env.get_scripture_text("elkton", ref)
        except Exception as e:
            logging.error("Elkton scripture fetch failed for '%s': %s", ref, e)
            elk_text = ""

        try:
            lb_text = openlp_env.get_scripture_text("lb", ref)
        except Exception as e:
            logging.error("LB scripture fetch failed for '%s': %s", ref, e)
            lb_text = ""

        if elk_text:
            md = replace_placeholder_with_block(md, elk_ph, elk_text)
        if lb_text:
            md = replace_placeholder_with_block(md, lb_ph, lb_text)

    return md


# ---------------------------------------------------------------------------
# Manual browser capture helpers (Offertory & Benediction)
# ---------------------------------------------------------------------------

def capture_clipboard(browser, label: str, url: str | None = None) -> str:
    """
    Open a new page in the given browser, optionally navigate to `url`,
    then let the user select and copy text (Ctrl+C). The function waits
    for the user to press Enter in the terminal and returns the clipboard
    contents as a string.
    """
    page = browser.new_page()
    print(f"\n=== {label} ===")
    if url:
        print(url)
        page.goto(url, wait_until="domcontentloaded")
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeoutError:
            pass
    else:
        print("A browser window has opened. Navigate to the site you want to copy from.")

    # Clear clipboard if possible
    if pyperclip is not None:
        try:
            pyperclip.copy("")
        except Exception:
            pass

    print("Select the text you want, press Ctrl+C, then press Enter here.")
    input("> ")

    text = ""
    if pyperclip is not None:
        try:
            text = (pyperclip.paste() or "").strip()
        except Exception as e:
            logging.error("Failed to read clipboard: %s", e)

    page.close()

    if not text:
        print(f"[Warning] No text captured for {label}.")
    else:
        print(f"[OK] Captured {len(text)} characters for {label}.")

    return text


def gather_offertory_and_benediction_via_browser(md: str, use_browser: bool) -> str:
    """
    If `use_browser` is True and browser dependencies are available, open a
    Firefox window and let the user copy Offertory and Benediction text into
    the clipboard, inserting the results into the master markdown.
    """
    if not use_browser:
        logging.info("Skipping browser-based capture (per --no-browser).")
        return md

    if not HAVE_BROWSER_DEPS:
        logging.error(
            "playwright/pyperclip not available; cannot do browser capture. "
            "Install 'playwright' (and run 'playwright install firefox') "
            "and 'pyperclip' if you want this feature."
        )
        return md

    need_offertory = "{offertory_text}" in md
    need_benediction = "{benediction}" in md

    if not (need_offertory or need_benediction):
        logging.info("No offertory/benediction placeholders found; skipping browser capture.")
        return md

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=False)
        try:
            if need_offertory:
                off_text = capture_clipboard(
                    browser,
                    "Offertory text (navigate to UMFMichigan / Discipleship or other source as desired)",
                    None,
                )
                if off_text:
                    md = append_below_placeholder(md, "offertory_text", off_text)

            if need_benediction:
                ben_text = capture_clipboard(
                    browser,
                    "Commission & Benediction (e.g., from Laughingbird)",
                    None,
                )
                if ben_text:
                    md = replace_placeholder_with_block(md, "benediction", ben_text)
        finally:
            browser.close()

    return md


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Weekly Automation: text gathering")
    parser.add_argument(
        "--master",
        required=True,
        help="Master markdown filename (relative to WORSHIP_DIR or absolute path)",
    )
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Do not open a browser for Offertory/Benediction; scripture only.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable debug logging.",
    )

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    master_path = Path(args.master)
    if not master_path.is_absolute():
        master_path = WORSHIP_DIR / master_path

    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    logging.info("Using master file: %s", master_path)

    md = read_text(master_path)

    # 1) Insert scripture blocks from OpenLP Bible databases
    md = insert_scripture_blocks(md)

    # 2) Optionally gather Offertory & Benediction text via browser
    md = gather_offertory_and_benediction_via_browser(md, use_browser=not args.no_browser)

    write_text(master_path, md)
    logging.info("Updated master file: %s", master_path)


if __name__ == "__main__":
    main()
