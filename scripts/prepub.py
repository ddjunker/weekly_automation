#!/usr/bin/env python3
"""
prepub.py — Pre-publication checks for text and music resources.

Reads the Master markdown and validates that all referenced resources
(scripture, CtW/AoF custom slides, songs) resolve in the local databases
without writing anything. Writes a report file for each section checked.

Usage:
    python scripts/prepub.py --master "Master 2025-11-23.md"        # both
    python scripts/prepub.py --master "Master 2025-11-23.md" -t     # text only
    python scripts/prepub.py --master "Master 2025-11-23.md" -m     # music only
"""

import argparse
import logging
import sys
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.utils.config import resolve_master_path
from scripts.utils.openlp import build_song_index
from scripts.text_gather import (
    _check_scripture,
    _check_custom_slides,
    _write_text_check_report,
)
from scripts.music_gather import (
    check_master,
    _write_music_check_report,
)


def _print_issues(results: list[dict], key: str) -> None:
    issues = [r for r in results
              if r[key] == "MISSING" or str(r[key]).startswith("MULTIPLE")]
    if issues:
        print(f"  Issues found: {len(issues)}")
        for r in issues:
            print(f"    [{r[key]}] {r}")
    else:
        print("  All OK.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Pre-publication resource checks")
    parser.add_argument("--master", required=True, help="Master markdown filename or path")
    parser.add_argument("-t", action="store_true",
                        help="Check text resources (scripture, CtW, AoF)")
    parser.add_argument("-m", action="store_true",
                        help="Check music resources (songs)")
    parser.add_argument("--verbose", action="store_true", help="Enable debug logging")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    run_all = not (args.t or args.m)
    run_t = run_all or args.t
    run_m = run_all or args.m

    master_path = resolve_master_path(args.master)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")

    md = master_path.read_text(encoding="utf-8")

    if run_t:
        results = _check_scripture(md) + _check_custom_slides(md)
        report_path = _write_text_check_report(master_path, results)
        print(f"\nText check report: {report_path}")
        _print_issues(results, key="status")

    if run_m:
        logging.info("Building OpenLP song indexes...")
        elk_index = build_song_index("elkton")
        lb_index = build_song_index("lb")
        results = check_master(md, elk_index, lb_index)
        report_path = _write_music_check_report(master_path, results)
        print(f"\nMusic check report: {report_path}")
        _print_issues(results, key="status")


if __name__ == "__main__":
    main()
