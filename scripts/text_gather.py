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
playwright_sync_api: Any = importlib.import_module("playwright.sync_api")
sync_playwright = playwright_sync_api.sync_playwright
PWTimeout = playwright_sync_api.TimeoutError

from scripts.utils.config import config, resolve_master_path
from scripts.utils.placeholder import (
    append_below_placeholder,
    extract_block,
    has_placeholder,
)
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
    m = re.match(r"^([1-3]?\s?[A-Za-z]+)\s+(\d+)(?::([\dA-Za-z\-,\s]+))?$", ref)
    if not m:
        return ref, "", ""  # let higher logic handle errors
    book = m.group(1)
    chapter = m.group(2)
    verses = m.group(3) or ""
    return book, chapter, verses



# ---------------- CtW + AoF Retrieval ----------------

def _append_slide_matches(md: str, church: str, placeholder_key: str, prefix: str, label: str, exact: bool = False) -> str:
    """
    Append all custom slides whose titles start with `prefix` (case-insensitive),
    inserting diagnostics into the markdown.
    When exact=True, only slides whose normalized title equals prefix are matched.
    """
    prefix_clean = clean_text(prefix).lower().replace(",", "")
    slides = list_custom_slides(church)  # [{uuid:int,title:str}, ...]

    if exact:
        matches = [s for s in slides if clean_text(s["title"]).lower().replace(",", "") == prefix_clean]
    else:
        matches = [s for s in slides if clean_text(s["title"]).lower().replace(",", "").startswith(prefix_clean)]

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


def gather_custom_slides(md: str, run_ctw: bool = True, run_aof: bool = True) -> str:
    # Inputs from markdown
    ctw_ref = extract_clean(md, "ctw_ref")      # e.g. "Psalm 122:1-9"
    aof_ref = extract_clean(md, "aof_ref")      # e.g. "39top" (NOT scripture)

    for church in ("elkton", "lb"):
        # CtW titles are like: "CtW Psalm 122:1-9 [Flag]"
        if run_ctw:
            ctw_ref_church = extract_clean(md, f"ctw_ref_{church}")
            if ctw_ref_church:
                md = _append_slide_matches(md, church, f"ctw_xml_{church}", ctw_ref_church, "CtW", exact=True)

        # AoF titles are like: "AoF p39top [Flag]"
        if run_aof and aof_ref:
            aof_prefix = f"AoF p{aof_ref}"
            md = _append_slide_matches(md, church, f"aof_xml_{church}", aof_prefix, "AoF")

            # Populate {aof_title} from the credits of the first Elkton match
            if church == "elkton" and has_placeholder(md, "aof_title"):
                prefix_clean = clean_text(aof_prefix).lower()
                matches = [s for s in list_custom_slides("elkton")
                           if clean_text(s["title"]).lower().startswith(prefix_clean)]
                if matches:
                    slide = load_custom_slide("elkton", matches[0]["uuid"])
                    credits = slide["credits"] if slide else ""
                    md = append_below_placeholder(md, "aof_title", credits if credits else f"[No credits for '{matches[0]['title']}']")
                else:
                    md = append_below_placeholder(md, "aof_title", f"[No AoF match found for '{aof_prefix}']")

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
                verse_start = re.match(r"\d+", verses)
                md = append_below_placeholder(
                    md, f"{key}_{church}_verses_start",
                    verse_start.group(0) if verse_start else verses,
                )
    return md



# ---------------- Dry-run checks ----------------

def _check_scripture(md: str) -> list[dict]:
    """Check all scripture references against the DB without writing. Returns list of result dicts."""
    results = []
    for idx in (1, 2, 3):
        ref = extract_block(md, f"scripture_ref_{idx}")
        if not ref:
            continue
        for church in ("elkton", "lb"):
            try:
                text = get_scripture_text(church, ref)
                verse_count = len([l for l in text.splitlines() if l.strip()])
                results.append({"section": "Scripture", "item": ref, "church": church,
                                "status": "found", "detail": f"{verse_count} verse(s)"})
            except Exception as e:
                results.append({"section": "Scripture", "item": ref, "church": church,
                                "status": "MISSING", "detail": str(e)})
    return results


def _check_custom_slides(md: str, check_ctw: bool = True, check_aof: bool = True) -> list[dict]:
    """Check CtW and AoF custom slide availability without writing. Returns list of result dicts."""
    results = []
    ctw_ref = extract_clean(md, "ctw_ref")
    aof_ref = extract_clean(md, "aof_ref")

    for church in ("elkton", "lb"):
        if check_ctw and ctw_ref:
            prefix = f"CtW {ctw_ref}"
            prefix_clean = clean_text(prefix).lower().replace(",", "")
            matches = [s for s in list_custom_slides(church)
                       if clean_text(s["title"]).lower().replace(",", "").startswith(prefix_clean)]
            if not matches:
                results.append({"section": "CtW", "item": prefix, "church": church,
                                "status": "MISSING", "detail": "No match found"})
            else:
                status = "found" if len(matches) == 1 else f"MULTIPLE ({len(matches)})"
                results.append({"section": "CtW", "item": prefix, "church": church,
                                "status": status, "detail": "; ".join(s["raw_title"] for s in matches)})

        if check_aof and aof_ref:
            prefix = f"AoF p{aof_ref}"
            prefix_clean = clean_text(prefix).lower()
            matches = [s for s in list_custom_slides(church)
                       if clean_text(s["title"]).lower().startswith(prefix_clean)]
            if not matches:
                results.append({"section": "AoF", "item": prefix, "church": church,
                                "status": "MISSING", "detail": "No match found"})
            else:
                status = "found" if len(matches) == 1 else f"MULTIPLE ({len(matches)})"
                results.append({"section": "AoF", "item": prefix, "church": church,
                                "status": status, "detail": "; ".join(s["raw_title"] for s in matches)})

    return results


def _write_text_check_report(master_path: Path, results: list[dict]) -> Path:
    """Write dry-run check results to a markdown report file next to the master."""
    report_path = master_path.with_name(master_path.stem + "_text_check.md")
    lines = [f"# Text Check: {master_path.name}\n"]

    sections = {}
    for r in results:
        sections.setdefault(r["section"], []).append(r)

    for section, rows in sections.items():
        lines.append(f"## {section}\n")
        lines.append("| Item | Church | Status | Detail |")
        lines.append("|------|--------|--------|--------|")
        for r in rows:
            lines.append(f"| {r['item']} | {r['church']} | {r['status']} | {r['detail']} |")
        lines.append("")

    if not results:
        lines.append("*(No checkable items found — offertory/benediction require interactive capture)*\n")

    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


# ---------------- Main ----------------

def main():
    parser = argparse.ArgumentParser(description="Weekly Automation Text Gatherer")
    parser.add_argument("--master", required=True)
    parser.add_argument("--open-output", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--dry-run", action="store_true",
                        help="Check DB lookups and report mismatches without writing to master")
    parser.add_argument("-s", action="store_true", help="Update scripture placeholders only")
    parser.add_argument("-c", action="store_true", help="Update CtW placeholders only")
    parser.add_argument("-a", action="store_true", help="Update AoF placeholders only")
    parser.add_argument("-w", action="store_true", help="Update offertory and benediction placeholders only")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s"
    )

    # If no section flags are given, run everything
    run_all = not any([args.s, args.c, args.a, args.w])
    run_s = run_all or args.s
    run_c = run_all or args.c
    run_a = run_all or args.a
    run_w = run_all or args.w

    master_path = resolve_master_path(args.master)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    md = master_path.read_text(encoding="utf-8")

    if args.dry_run:
        results = []
        if run_s:
            results.extend(_check_scripture(md))
        if run_c or run_a:
            results.extend(_check_custom_slides(md, check_ctw=run_c, check_aof=run_a))
        report_path = _write_text_check_report(master_path, results)
        print(f"\nDry-run report written: {report_path}\n")
        missing = [r for r in results if r["status"].startswith("MISSING") or r["status"].startswith("MULTIPLE")]
        if missing:
            print(f"Issues found: {len(missing)}")
            for r in missing:
                print(f"  [{r['status']}] {r['section']} / {r['item']} ({r['church']}): {r['detail']}")
        else:
            print("All checked items resolved successfully.")
        return

    if run_w:
        if check_firefox_running():
            print("Closing existing Firefox instances...")
            close_firefox_instances()

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

            off1_url = extract_clean(md, "offertory_source")
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
            ben1 = maybe_capture(page, "Commission & Benediction", ben1_url)
            ben2 = maybe_capture(page, "Commission & Benediction (Alt)", ben2_url)

            context.close()

        if off1:
            md = append_below_placeholder(md, "offertory_source_text", off1)
        if off2:
            md = append_below_placeholder(md, "offertory_source_alt_text", off2)
        if ben1:
            md = append_below_placeholder(md, "benediction_text", ben1)
        if ben2:
            md = append_below_placeholder(md, "benediction_text_alt", ben2)

    if run_s:
        md = gather_scripture_text(md)

    if run_c or run_a:
        md = gather_custom_slides(md, run_ctw=run_c, run_aof=run_a)

    master_path.write_text(md, encoding="utf-8")
    print(f"\nUpdated: {master_path}\n")
    if args.open_output:
        import os, sys
        if sys.platform.startswith("win"):
            os.startfile(master_path)


if __name__ == "__main__":
    main()
