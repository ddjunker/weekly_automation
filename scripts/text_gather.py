#!/usr/bin/env python3
"""
text_gather.py — Weekly Automation: Scripture + Offertory + CtW + AoF Gatherer
"""

import subprocess
import importlib
import psutil
import argparse
import logging
import re
import time
import platform
from pathlib import Path
from typing import Any
import yaml

playwright_sync_api: Any = importlib.import_module("playwright.sync_api")
sync_playwright = playwright_sync_api.sync_playwright
PWTimeout = playwright_sync_api.TimeoutError

from scripts.utils.config import config
from scripts.utils.placeholder import append_below_placeholder, extract_block
from scripts.utils.text_clean import clean_markdown, clean_text, xml_to_text
from scripts.utils.openlp import (
    get_scripture_text,
    list_custom_slides,
    load_custom_slide
)


# ---------------- this is to avoid the empty clipboard problem encountered with one of the sites -----------------

def check_firefox_running():
    """Check if any Firefox processes are running."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'firefox.exe':  # Use 'firefox' on Linux/Mac
            return True
    return False


def close_firefox_instances():
    """Close all running Firefox processes."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'firefox.exe':  # Use 'firefox' on Linux/Mac
            try:
                proc.terminate()  # Gracefully terminate the process
                proc.wait(timeout=5)  # Wait for the process to terminate
            except (psutil.AccessDenied, psutil.NoSuchProcess, psutil.TimeoutExpired) as e:
                logging.warning("Could not terminate Firefox process pid=%s: %s", proc.pid, e)


# ----- determines where to look ---------

def resolve_master_path(master_arg: str) -> Path:
    """
    Resolve the Master markdown file path:

    1. If master_arg is an absolute path → return it unchanged.
    2. Otherwise → return config.worship_dir / master_arg
    3. worship_dir must exist (strict behavior).

    Returns:
        Path object to the master file.
    """

    mp = Path(master_arg)

    # Case 1: absolute override
    if mp.is_absolute():
        return mp

    # Case 2: relative → use worship_dir
    worship = config.worship_dir.expanduser().resolve()
    if not worship.exists():
        raise FileNotFoundError(f"Worship directory not found: {worship}")

    return worship / master_arg


# ---------------- Clipboard Capture ----------------

def _read_system_clipboard() -> str:
    """Read text from the OS clipboard (preferred over browser clipboard API)."""
    system = platform.system().lower()

    try:
        if system == "windows":
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"],
                capture_output=True,
                text=True,
                check=False,
            )
            return (result.stdout or "").strip()

        if system == "darwin":
            result = subprocess.run(["pbpaste"], capture_output=True, text=True, check=False)
            return (result.stdout or "").strip()

        # Linux fallback (xclip)
        result = subprocess.run(
            ["xclip", "-selection", "clipboard", "-o"],
            capture_output=True,
            text=True,
            check=False,
        )
        return (result.stdout or "").strip()
    except Exception:
        return ""

def capture_clipboard(page, label: str, url: str) -> str:
    print(f"\n=== {label} ===")
    print(url)
    try:
        page.goto(url, wait_until="domcontentloaded")
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeout:
        pass

    print("Select the exact text, press Ctrl+C, then press Enter here.")
    input()

    text = ""

    # Prefer OS clipboard reads (more reliable than browser clipboard context).
    for _ in range(8):
        text = _read_system_clipboard()
        if text:
            break
        time.sleep(0.2)

    # Fallback: browser clipboard API
    if not text:
        try:
            text = page.evaluate("async () => await navigator.clipboard.readText()")
        except Exception:
            text = ""

    if not isinstance(text, str):
        text = ""

    text = text.strip()
    print(f"Captured {len(text)} characters.")
    return text


# ---------------- CtW Support ----------------

# def parse_ctw_title(title: str):
#     title = title.strip()
#     if " UMH" in title:
#         base, flag = title.split(" UMH", 1)
#         return base.strip(), ("UMH" + flag.strip())
#     return title, None


def extract_clean(md: str, key: str) -> str:
    raw = extract_block(md, key)

    if isinstance(raw, (list, tuple)):
        parts = raw
    else:
        parts = [raw]

    parts = [
        str(p).strip()
        for p in parts
        if p and str(p).strip() and not str(p).strip().startswith("%%")
    ]

    flat = " ".join(parts)
    return clean_text(flat)


def normalize_reference(ref: str):
    """
    Input like 'Psalm 122:1-9'
    Returns: (book='Psalm', chapter='122', verses='1-9' or '')
    """
    ref = ref.replace("–", "-").replace("—", "-").strip()
    m = re.match(r"^([1-3]?\s?[A-Za-z]+)\s+(\d+)(?::([\d\-]+))?$", ref)
    if not m:
        return ref, "", ""  # let higher logic handle errors
    book = m.group(1)
    chapter = m.group(2)
    verses = m.group(3) or ""
    return book, chapter, verses



# ---------------- CtW + AoF Retrieval ----------------

def _append_slide_matches(md: str, church: str, placeholder_key: str, prefix: str, label: str) -> str:
    """
    Append all custom slides whose titles start with `prefix` (case-insensitive),
    inserting diagnostics into the markdown.
    """
    prefix_clean = clean_text(prefix).lower()
    slides = list_custom_slides(church)  # [{uuid:int,title:str}, ...]

    matches = [s for s in slides if clean_text(s["title"]).lower().startswith(prefix_clean)]

    if not matches:
        return append_below_placeholder(md, placeholder_key, f"[No {label} match found for '{prefix}' in {church}]")

    # Diagnostics: if multiple variants exist (differ only by flags), say so
    if len(matches) > 1:
        md = append_below_placeholder(md, placeholder_key, f"[{label} multiple variants: {len(matches)}]")

    for s in matches:
        slide = load_custom_slide(church, s["uuid"])   # dict {id,title,text}
        if not slide or not isinstance(slide.get("text"), str) or not slide["text"].strip():
            md = append_below_placeholder(md, placeholder_key, f"[{label} slide {s['uuid']} had no text]")
            continue
        cleaned_text = xml_to_text(slide["text"])
        # Remove OpenLP-specific formatting tags
        cleaned_text = re.sub(r'\[={3,}\]|\[[\-—]{3,}\]|\{y\}|\{/y\}', '', cleaned_text)
        md = append_below_placeholder(md, placeholder_key, cleaned_text)

    return md


def gather_custom_slides(md: str) -> str:
    # Inputs from markdown
    ctw_ref = extract_clean(md, "ctw_ref")      # e.g. "Psalm 122:1-9"
    aof_ref = extract_clean(md, "aof_ref")      # e.g. "39top" (NOT scripture)

    for church in ("elkton", "lb"):
        # CtW titles are like: "CtW Psalm 122:1-9 [Flag]"
        if ctw_ref:
            ctw_prefix = f"CtW {ctw_ref}"
            md = _append_slide_matches(md, church, f"ctw_xml_{church}", ctw_prefix, "CtW")

        # AoF titles are like: "AoF p39top [Flag]"
        if aof_ref:
            aof_prefix = f"AoF p{aof_ref}"
            md = _append_slide_matches(md, church, f"aof_xml_{church}", aof_prefix, "AoF")

    return md



# ---------------- Scripture Retrieval ----------------

def gather_scripture_text(md: str) -> str:
    for idx in (1, 2, 3):
        key = f"scripture_ref_{idx}"
        ref = extract_block(md, key)
        if not ref:
            continue

        for church in ("elkton", "lb"):
            try:
                text = get_scripture_text(church, ref)
                md = append_below_placeholder(md, f"{key}_{church}", text)
            except Exception as e:
                md = append_below_placeholder(
                    md, f"{key}_{church}",
                    f"[Scripture fetch error for {ref} in {church}: {e}]"
                )
            # Only for idx=3 and church=elkton, normalize reference
            if idx == 3 and church == "elkton":
                book, chapter, verses = normalize_reference(ref)
                md = append_below_placeholder(md, f"{key}_{church}_book", book)
                md = append_below_placeholder(md, f"{key}_{church}_chapter", chapter)
                md = append_below_placeholder(md, f"{key}_{church}_verses", verses)
    return md


# ---------------- Service Info Block ----------------

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
    else:
        # Append at end if not found
        return md.rstrip() + "\n\n" + new_block + "\n"


# ---------------- Main ----------------

def main():
    parser = argparse.ArgumentParser(description="Weekly Automation Text Gatherer")
    parser.add_argument("--master", required=True)
    parser.add_argument("--open-output", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s"
    )



    if check_firefox_running():
        print("Closing existing Firefox instances...")
        close_firefox_instances()


    master_path = resolve_master_path(args.master)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    md = master_path.read_text(encoding="utf-8")

    with sync_playwright() as p:
        context = p.firefox.launch_persistent_context(
            user_data_dir=str(config.browser_profile),
            headless=False,
            firefox_user_prefs={
                "dom.events.asyncClipboard.readText": True,
                "dom.events.testing.asyncClipboard": True,
                "permissions.default.clipboard-read": 1,
                "permissions.default.clipboard-write": 1,
            }
        )
        page = context.new_page()

        # Extract URLs from Markdown
        off1_url_raw = extract_block(md, "offertory_source")
        off1_url = extract_clean(md, "offertory_source")
        off2_url_raw = extract_block(md, "offertory_source_alt")
        off2_url = extract_clean(md, "offertory_source_alt")
        ben1_url = extract_clean(md, "benediction_source")
        ben2_url = extract_clean(md, "benediction_source_alt")

        def maybe_capture(page, label, url):
            if not url or url.lower() == "na":
                print(f"\n=== {label} ===\nSkipped (source='{url}')")
                return ""
            if not url.lower().startswith("http"):
                print(f"\n=== {label} ===\nSkipped (invalid URL='{url}')")
                return ""
            return capture_clipboard(page, label, url)

        off1 = maybe_capture(page, "Offertory Prayer (Primary)", off1_url)
        off2 = maybe_capture(page, "Offertory Prayer (Alternate)", off2_url)
        ben1  = maybe_capture(page, "Commission & Benediction", ben1_url)
        ben2  = maybe_capture(page, "Commission & Benediction (Alt)", ben2_url)

        context.close()

    # Insert into markdown
    # Offertory
    if off1:
        md = append_below_placeholder(md, "offertory_source_text", off1)
    if off2:
        md = append_below_placeholder(md, "offertory_source_alt_text", off2)
    # Benediction
    if ben1:
        md = append_below_placeholder(md, "benediction_text", ben1)
    if ben2:
        md = append_below_placeholder(md, "benediction_text_alt", ben2)
    md = gather_scripture_text(md)
    md = gather_custom_slides(md)

    # Update serviceinfo block with ctw and scriptures
    ctw_val = extract_clean(md, "ctw_ref")
    scriptures = []
    for idx in (1, 2, 3):
        ref = extract_clean(md, f"scripture_ref_{idx}")
        if ref:
            scriptures.append(ref)
    updates = {}
    if ctw_val:
        updates["ctw"] = ctw_val
    if scriptures:
        updates["scriptures"] = scriptures
    if updates:
        md = update_serviceinfo_block(md, updates)

    master_path.write_text(md, encoding="utf-8")
    print(f"\nUpdated: {master_path}\n")
    if args.open_output:
        import os, sys
        if sys.platform.startswith("win"):
            os.startfile(master_path)


if __name__ == "__main__":
    main()
