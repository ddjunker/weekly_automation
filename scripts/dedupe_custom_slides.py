"""
dedupe_custom_slides.py — one-off cleanup for OpenLP custom slide libraries.

OpenLP re-links a custom service item to a library slide by title + credits
(custom plugin service_load). When that lookup fails and "add custom slides from
service" is enabled, OpenLP silently inserts a copy of the service item on every
load — which is how duplicate rows such as six 'Offertory Prayer' entries built
up. publish.py no longer produces unmatchable items; this script prunes the rows
that accumulated before that fix.

Rows are grouped by exact (title, credits, theme_name). One row per group is
kept — the one OpenLP itself resolves to, i.e. the lowest id — and the rest are
deleted. Dry run by default; --apply writes, after taking a timestamped backup
of the database next to it.

Usage:
    python -m scripts.dedupe_custom_slides                      # dry run, both churches
    python -m scripts.dedupe_custom_slides --title "Offertory"  # only matching titles
    python -m scripts.dedupe_custom_slides --apply
"""

from __future__ import annotations

import argparse
import shutil
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.utils.openlp import get_custom_db

CHURCHES = ("elkton", "lb")


def _preview(xml_text: str, width: int = 70) -> str:
    """One-line preview of a slide's stored text."""
    from scripts.utils.openlp import extract_openlp_slide_text

    text = " ".join(extract_openlp_slide_text(xml_text or "").split())
    return text[:width] + ("…" if len(text) > width else "")


def _duplicate_groups(conn: sqlite3.Connection, title_filter: str) -> list[list[sqlite3.Row]]:
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT id, title, IFNULL(credits, '') AS credits, IFNULL(theme_name, '') AS theme_name, text "
        "FROM custom_slide ORDER BY id"
    ).fetchall()

    groups: dict[tuple[str, str, str], list[sqlite3.Row]] = {}
    needle = title_filter.strip().lower()
    for row in rows:
        if needle and needle not in (row["title"] or "").lower():
            continue
        groups.setdefault((row["title"], row["credits"], row["theme_name"]), []).append(row)

    return [g for g in groups.values() if len(g) > 1]


def dedupe_church(church: str, *, title_filter: str, apply: bool) -> int:
    """Report (and optionally delete) duplicate rows. Returns rows deleted."""
    db_path = get_custom_db(church)
    if not db_path.exists():
        print(f"[{church}] custom slide DB not found: {db_path}")
        return 0

    conn = sqlite3.connect(db_path, timeout=10)
    try:
        groups = _duplicate_groups(conn, title_filter)
        if not groups:
            print(f"[{church}] no duplicate title+credits groups found")
            return 0

        doomed: list[int] = []
        for group in groups:
            keep, extras = group[0], group[1:]
            print(f"\n[{church}] {keep['title']!r} | credits={keep['credits']!r}")
            print(f"    keep   id={keep['id']}: {_preview(keep['text'])}")
            for row in extras:
                print(f"    delete id={row['id']}: {_preview(row['text'])}")
                doomed.append(int(row["id"]))

        if not apply:
            print(f"\n[{church}] dry run — {len(doomed)} row(s) would be deleted. Re-run with --apply.")
            return 0

        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = db_path.with_name(f"{db_path.stem}.dedupe-backup-{stamp}{db_path.suffix}")
        shutil.copy2(db_path, backup)
        print(f"\n[{church}] backup written: {backup}")

        conn.executemany("DELETE FROM custom_slide WHERE id = ?", [(i,) for i in doomed])
        conn.commit()
        print(f"[{church}] deleted {len(doomed)} row(s)")
        return len(doomed)
    finally:
        conn.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Remove duplicate OpenLP custom slides.")
    parser.add_argument("--church", choices=[*CHURCHES, "both"], default="both")
    parser.add_argument("--title", default="", help="only consider titles containing this text")
    parser.add_argument("--apply", action="store_true", help="delete rows (default is a dry run)")
    args = parser.parse_args()

    churches = CHURCHES if args.church == "both" else (args.church,)
    total = 0
    for church in churches:
        total += dedupe_church(church, title_filter=args.title, apply=args.apply)

    if args.apply:
        print(f"\nDone — {total} row(s) deleted. Close OpenLP before reopening the library.")


if __name__ == "__main__":
    main()
