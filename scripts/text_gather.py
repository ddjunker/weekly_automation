#!/usr/bin/env python3
"""
text_gather.py — Weekly Automation: Scripture + Offertory + CtW + AoF Gatherer
"""

import argparse
import logging
import re
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from scripts.utils.config import config
from scripts.utils.placeholder import append_below_placeholder, extract_block
from scripts.utils.text_clean import clean_markdown, clean_text
from scripts.utils.openlp import (
    get_scripture_text,
    list_custom_slides,
    load_custom_slide
)




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

    try:
        text = page.evaluate("navigator.clipboard.readText()")
    except Exception:
        text = ""
    text = (text or "").strip()
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



# def ctw_matches(slide_title: str, ref_title: str) -> bool:
#     s = slide_title.lower()
#     r = ref_title.lower()

#     if s == r:
#         return True

#     pat = r"([1-3]?\s*[a-zA-Z]+)\s+(\d+)"
#     m1 = re.search(pat, s)
#     m2 = re.search(pat, r)
#     if m1 and m2 and m1.group(1) == m2.group(1) and m1.group(2) == m2.group(2):
#         return True

#     return False


# ---------------- CtW + AoF Retrieval ----------------

def gather_custom_slides(md: str) -> str:
        # --- Extract & clean CtW reference ---
    ctw_ref = extract_clean(md, "ctw_ref")
    aof_ref = extract_clean(md, "aof_ref")

    for church in ("elkton", "lb"):
        db_path = (
            config.elkton_root / "custom" / "custom.sqlite"
            if church == "elkton"
            else config.lb_root / "custom" / "custom.sqlite"
        )

        # -------------------------------------------------------
        # Call to Worship (CtW)
        # -------------------------------------------------------
        ctw_ref = extract_clean(md, "ctw_ref")

        if ctw_ref:
            try:
                book, chapter, verses = normalize_reference(ctw_ref)
                ref_exact = f"{book} {chapter}" + (f":{verses}" if verses else "")
            except Exception as e:
                md = append_below_placeholder(md, "ctw_xml_elkton", f"[CtW parse error: {e}]")
                md = append_below_placeholder(md, "ctw_xml_lb", f"[CtW parse error: {e}]")
                return md

            ref_exact_lower = f"CtW {ref_exact}".lower()
            ref_chapter_lower = f"CtW {book} {chapter}".lower()

            for church in ("elkton", "lb"):
                db_path = (
                    config.elkton_root / "custom" / "custom.sqlite"
                    if church == "elkton"
                    else config.lb_root / "custom" / "custom.sqlite"
                )
                all_slides = list_custom_slides(church)

                # --- Primary exact matches ---
                exact_matches = [
                    s for s in all_slides
                    if s["title"].lower().startswith(ref_exact_lower)
                ]

                if exact_matches:
                    # Insert a header so the user can see what matched
                    md = append_below_placeholder(
                        md,
                        f"ctw_xml_{church}",
                        f"[CtW exact match: {ref_exact}]"
                    )

                    for s in exact_matches:
                        uuid = s["uuid"]
                        slide = load_custom_slide(church, uuid)
                        if slide:
                            md = append_below_placeholder(md, f"ctw_xml_{church}", slide["text"])  
                    continue  # ***Do NOT fall back if exact matches exist***

                # --- Fallback: chapter-only (Psalm 122) ---
                chapter_matches = [
                    s for s in all_slides
                    if s["title"].lower().startswith(ref_chapter_lower)
                ]

                if chapter_matches:
                    md = append_below_placeholder(
                        md,
                        f"ctw_xml_{church}",
                        f"[CtW fallback: {book} {chapter}]"
                    )
                    for s in chapter_matches:
                        uuid = s["uuid"]
                        slide = load_custom_slide(church, uuid)
                        if slide:
                            md = append_below_placeholder(md, f"ctw_xml_{church}", slide["text"])  
                else:
                    md = append_below_placeholder(
                        md,
                        f"ctw_xml_{church}",
                        f"[No CtW match found for {ref_exact}]"
                    )


        # --- Affirmation of Faith (AoF) ---
        aof_ref = extract_clean(md, "aof_ref")

        if aof_ref:
            try:
                a_book, a_chapter, a_verses = normalize_reference(aof_ref)
                a_ref_exact = f"{a_book} {a_chapter}" + (f":{a_verses}" if a_verses else "")
            except Exception as e:
                for church in ("elkton", "lb"):
                    md = append_below_placeholder(
                        md,
                        f"aof_xml_{church}",
                        f"[AoF parse error: {e}]"
                    )
                # skip further AoF work if reference is invalid
                pass
            else:
                # Lowercase forms for matching slide titles
                a_ref_exact_lower = f"AoF {a_ref_exact}".lower()
                a_ref_ch_lower = f"AoF {a_book} {a_chapter}".lower()

                for church in ("elkton", "lb"):
                    db_path = (
                        config.elkton_root / "custom" / "custom.sqlite"
                        if church == "elkton"
                        else config.lb_root / "custom" / "custom.sqlite"
                    )
                    all_slides = list_custom_slides(church)

                    # --- Primary exact lookup ---
                    exact_matches = [
                        s for s in all_slides
                        if s["title"].lower().startswith(a_ref_exact_lower)
                    ]

                    if exact_matches:
                        md = append_below_placeholder(
                            md,
                            f"aof_xml_{church}",
                            f"[AoF exact match: {a_ref_exact}]"
                        )
                        for s in exact_matches:
                            uuid = s["uuid"]
                            slide = load_custom_slide(church, uuid)
                            if slide:
                                md = append_below_placeholder(md, f"aof_xml_{church}", slide["text"])  
                        continue  # Do NOT fallback if exact matches exist

                    # --- Fallback: Book + Chapter only ---
                    fallback_matches = [
                        s for s in all_slides
                        if s["title"].lower().startswith(a_ref_ch_lower)
                    ]

                    if fallback_matches:
                        md = append_below_placeholder(
                            md,
                            f"aof_xml_{church}",
                            f"[AoF fallback: {a_book} {a_chapter}]"
                        )
                        for s in fallback_matches:
                            uuid = s["uuid"]
                            slide = load_custom_slide(church, uuid)
                            if slide:
                                md = append_below_placeholder(md, f"aof_xml_{church}", slide["text"])  
                    else:
                        md = append_below_placeholder(
                            md,
                            f"aof_xml_{church}",
                            f"[No AoF match found for {a_ref_exact}]"
                        )


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
    return md


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

    master_path = resolve_master_path(args.master)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    md = master_path.read_text(encoding="utf-8")

    with sync_playwright() as p:
        context = p.firefox.launch_persistent_context(
            user_data_dir=str(config.browser_profile),
            headless=False
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

    master_path.write_text(md, encoding="utf-8")
    print(f"\nUpdated: {master_path}\n")

    if args.open_output:
        import os, sys
        if sys.platform.startswith("win"):
            os.startfile(master_path)


if __name__ == "__main__":
    main()
