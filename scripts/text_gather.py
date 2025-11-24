"""
text_gather.py – Weekly Info Gathering (Final Version, Refactored)
------------------------------------------------------------------
Uses flexible source URLs embedded in the Master markdown file.
Resolves Master file paths through weekly_config.ini (via config module).
Fetches Scripture from OpenLP Bible databases.
Captures Offertory/Benediction text through Firefox + clipboard.

Now refactored to use:
  • scripts.utils.placeholder.append_below_placeholder
  • scripts.utils.text_clean.clean_markdown
"""

import argparse
import logging
import sys
import re
from pathlib import Path

# Project utilities
from scripts.utils import openlp_env
from scripts.utils.config import config
from scripts.utils.placeholder import append_below_placeholder
from scripts.utils.text_clean import clean_markdown


# =========================================================
# BROWSER SUPPORT
# =========================================================
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    import pyperclip
    HAVE_BROWSER_DEPS = True
except ImportError:
    HAVE_BROWSER_DEPS = False


# =========================================================
# CONFIG + MASTER PATH RESOLUTION
# =========================================================

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

    # Case 1: absolute path override
    if mp.is_absolute():
        return mp

    # Case 2: relative → use worship_dir
    worship = config.worship_dir.expanduser().resolve()
    if not worship.exists():
        raise FileNotFoundError(f"Worship directory not found: {worship}")

    return worship / master_arg


# =========================================================
# MARKDOWN HELPERS
# =========================================================

def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


def get_source_url(md: str, key: str) -> str | None:
    """
    Read URL immediately following placeholder {key}.
    Returns None if missing or equal to 'na'.
    """
    lines = md.splitlines()
    for i, line in enumerate(lines):
        if line.strip() == f"{{{key}}}":
            if i + 1 < len(lines):
                candidate = lines[i + 1].strip()
                if candidate.lower().startswith("http") and candidate.lower() != "na":
                    return candidate
    return None


# =========================================================
# BROWSER CLIPBOARD GATHERING
# =========================================================

def capture_clipboard(browser, label: str, url: str) -> str:
    """
    Load URL in a new browser tab, user selects text & copies with Ctrl+C,
    then press Enter in the terminal. Clipboard text returned.
    """
    page = browser.new_page()
    print(f"\n=== {label} ===\n{url}")

    try:
        page.goto(url, wait_until="domcontentloaded")
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PWTimeout:
            # Not fatal; we just continue with whatever loaded.
            pass
    except Exception as e:
        print(f"[Browser Error] Unable to open URL: {e}")
        page.close()
        return ""

    try:
        pyperclip.copy("")
    except Exception:
        pass

    print("Select text on page, press Ctrl+C, then press Enter here.")
    input()

    text = (pyperclip.paste() or "").strip()
    page.close()

    if text:
        print(f"Captured {len(text)} characters.")
    else:
        print("[Warning] No text captured.")
    return text


def gather_offertory_and_benediction(md: str, use_browser: bool) -> str:
    """
    Reads Offertory and Benediction URLs from the Master markdown
    and performs browser-based clipboard capture.
    """
    if not use_browser:
        logging.info("Skipping browser capture (--no-browser).")
        return md

    if not HAVE_BROWSER_DEPS:
        logging.error("Missing playwright/pyperclip; skipping browser capture.")
        return md

    offertory_url    = get_source_url(md, "offertory_source")
    offertory_alt    = get_source_url(md, "offertory_source_alt")
    benediction_url  = get_source_url(md, "benediction_source")
    benediction_alt  = get_source_url(md, "benediction_source_alt")

    with sync_playwright() as p:
        profile_dir = Path.home() / ".weekly_info_profile"
        browser = p.firefox.launch_persistent_context(str(profile_dir), headless=False)

        try:
            if offertory_url:
                txt = capture_clipboard(browser, "Offertory (Primary)", offertory_url)
                if txt:
                    txt = clean_markdown(txt)
                    md = append_below_placeholder(md, "offertory_source_text", txt)

            if offertory_alt:
                txt = capture_clipboard(browser, "Offertory (Alternate)", offertory_alt)
                if txt:
                    txt = clean_markdown(txt)
                    md = append_below_placeholder(md, "offertory_source_alt_text", txt)

            if benediction_url:
                txt = capture_clipboard(browser, "Benediction (Primary)", benediction_url)
                if txt:
                    txt = clean_markdown(txt)
                    md = append_below_placeholder(md, "benediction_text", txt)

            if benediction_alt:
                txt = capture_clipboard(browser, "Benediction (Alternate)", benediction_alt)
                if txt:
                    txt = clean_markdown(txt)
                    md = append_below_placeholder(md, "benediction_text_alt", txt)

        finally:
            browser.close()

    return md


# =========================================================
# SCRIPTURE GATHERING
# =========================================================

def extract_scripture_refs(md: str):
    """
    Returns list of (placeholder, reference) extracted from markdown.
    """
    result = []
    pattern = re.compile(r"\{(scripture_ref_[^}]+)\}")
    lines = md.splitlines()

    for i, line in enumerate(lines):
        m = pattern.search(line)
        if m:
            key = m.group(1)
            ref_line = ""
            for j in range(i + 1, len(lines)):
                if lines[j].strip():
                    ref_line = lines[j].strip()
                    break
            if ref_line:
                result.append((key, ref_line))

    return result


def gather_scripture_text(md: str) -> str:
    refs = extract_scripture_refs(md)
    for key, ref in refs:
        for church in ["elkton", "lb"]:
            try:
                text = openlp_env.get_scripture_text(church, ref)
            except Exception as e:
                text = f"[Scripture error: {e}]"

            md = append_below_placeholder(md, f"{key}_{church}", text)

    return md


from scripts.utils.openlp_helpers import load_openlp_custom_slide
from scripts.utils.placeholder import append_below_placeholder, replace_placeholder
from scripts.utils.text_clean import clean_markdown


def gather_custom_slides(md: str) -> str:
    """
    Fetch Call to Worship (CtW) and Affirmation of Faith (AoF)
    from each church's OpenLP custom slide database.

    Requirements:
      • CtW lookup uses placeholder: {ctw_title}
      • AoF lookup uses placeholder: {aof}
      • Text output goes below:
            {ctw_text_elkton}, {ctw_text_lb},
            {aof_text_elkton}, {aof_text_lb}
      • UUID output goes below:
            {ctw_uuid_elkton}, {ctw_uuid_lb},
            {aof_uuid_elkton}, {aof_uuid_lb}
    """
    # read CtW + AoF titles
    ctw_ref = extract_block(md, "ctw_title")
    aof_ref = extract_block(md, "aof")

    # "CtW Psalm 145"
    def normalize_ctw(ref: str | None):
        if not ref:
            return None
        r = ref.strip()
        # accept “Psalm 145:1–5” but CtW slides usually only include chapter
        r = re.sub(r"[:].*$", "", r)  # drop verse range
        return f"CtW {r}".strip()

    ctw_title = normalize_ctw(ctw_ref)
    aof_title = aof_ref.strip() if aof_ref else None

    for church in ("elkton", "lb"):
        if church == "elkton":
            db_path = config.elkton_root / "custom" / "custom.sqlite"
        else:
            db_path = config.lb_root / "custom" / "custom.sqlite"

        # CtW slide fetch
        if ctw_title:
            try:
                slide = load_openlp_custom_slide(db_path, title=ctw_title)
                if slide:
                    uuid = slide["uuid"]
                    text = clean_markdown(slide["text"])
                    md = append_below_placeholder(md, f"ctw_uuid_{church}", str(uuid))
                    md = append_below_placeholder(md, f"ctw_text_{church}", text)
                else:
                    md = append_below_placeholder(
                        md,
                        f"ctw_text_{church}",
                        f"[No CtW slide found for {ctw_title} in {church}]"
                    )
            except Exception as e:
                md = append_below_placeholder(
                    md,
                    f"ctw_text_{church}",
                    f"[Error loading CtW for {ctw_title} in {church}: {e}]"
                )

        # AoF slide fetch
        if aof_title:
            try:
                slide = load_openlp_custom_slide(db_path, title=aof_title)
                if slide:
                    uuid = slide["uuid"]
                    text = clean_markdown(slide["text"])
                    md = append_below_placeholder(md, f"aof_uuid_{church}", str(uuid))
                    md = append_below_placeholder(md, f"aof_text_{church}", text)
                else:
                    md = append_below_placeholder(
                        md,
                        f"aof_text_{church}",
                        f"[No AoF slide found for {aof_title} in {church}]"
                    )
            except Exception as e:
                md = append_below_placeholder(
                    md,
                    f"aof_text_{church}",
                    f"[Error loading AoF for {aof_title} in {church}: {e}]"
                )

    return md


# =========================================================
# MAIN
# =========================================================

def main():
    parser = argparse.ArgumentParser(description="Weekly Info Gathering")
    parser.add_argument(
        "--master",
        required=True,
        help="Master filename (relative to worship_dir) or absolute path."
    )
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Skip browser Offertory/Benediction capture."
    )
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s"
    )

    master_path = resolve_master_path(args.master)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    md = read_text(master_path)

    # Step 1: Offertory/Benediction (browser scraping)
    md = gather_offertory_and_benediction(md, use_browser=not args.no_browser)

    # Step 2: Scripture (OpenLP)
    md = gather_scripture_text(md)
    
    # Step 3: CtW and Aof from OpenLP
    md = gather_custom_slides(md)

    write_text(master_path, md)
    print(f"\n✅ Updated: {master_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nAborted by user.")
        sys.exit(1)
