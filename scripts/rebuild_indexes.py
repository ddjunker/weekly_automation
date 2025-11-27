"""
Rebuild OpenLP index markdown files for Weekly Automation.

Supports:
    --slides   Rebuild slide indexes
    --songs    Rebuild song indexes
    --all      Rebuild everything (default)

Church selection:
    --elkton   Only Elkton
    --lb       Only LB
    (default = both)

Usage examples:
    python -m scripts.rebuild_indexes
    python -m scripts.rebuild_indexes --slides --elkton
    python -m scripts.rebuild_indexes --songs --lb
    python -m scripts.rebuild_indexes --all
"""

import argparse
from pathlib import Path

from scripts.utils.config import config
from scripts.utils.openlp_helpers import rebuild_custom_index

# ---------------------------------------------------------
# Placeholder for future song-index builder
# ---------------------------------------------------------
def rebuild_song_index(db_path: Path, out_path: Path):
    """
    Placeholder — implement once your song index format is finalized.
    This function is intentionally simple for now.
    """
    if not db_path.exists():
        print(f"[WARN] Song DB not found: {db_path}")
        return

    out_path.write_text(
        "# Song index generation is not yet implemented.\n",
        encoding="utf-8"
    )
    print(f"[OK] Created placeholder song index → {out_path}")


# ---------------------------------------------------------
# Main rebuild function
# ---------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Rebuild OpenLP index files")

    group = parser.add_mutually_exclusive_group()
    group.add_argument("--slides", action="store_true", help="Rebuild slide indexes only")
    group.add_argument("--songs", action="store_true", help="Rebuild song indexes only")
    group.add_argument("--all", action="store_true", help="Rebuild slides and songs (default)")

    parser.add_argument("--elkton", action="store_true", help="Only rebuild Elkton indexes")
    parser.add_argument("--lb", action="store_true", help="Only rebuild LB indexes")

    args = parser.parse_args()

    # What to rebuild
    do_slides = args.slides or args.all or not (args.slides or args.songs or args.all)
    do_songs  = args.songs  or args.all or not (args.slides or args.songs or args.all)

    # Which churches
    do_elkton = args.elkton or not (args.elkton or args.lb)
    do_lb     = args.lb     or not (args.elkton or args.lb)

    print("=== Rebuilding Indexes ===\n")

    if do_slides:
        print("→ Rebuilding Slide Indexes")
        if do_elkton:
            db = config.elkton_root / "custom" / "custom.sqlite"
            out = config.worship_dir / "openlp_custom_index_elkton.md"
            if db.exists():
                rebuild_custom_index(db, out)
                print(f"[OK] Elkton slides → {out}")
            else:
                print(f"[WARN] Elkton slide DB missing → {db}")

        if do_lb:
            db = config.lb_root / "custom" / "custom.sqlite"
            out = config.worship_dir / "openlp_custom_index_lb.md"
            if db.exists():
                rebuild_custom_index(db, out)
                print(f"[OK] LB slides → {out}")
            else:
                print(f"[WARN] LB slide DB missing → {db}")

        print()

    if do_songs:
        print("→ Rebuilding Song Indexes")
        if do_elkton:
            db = config.elkton_root / "songs" / "songs.sqlite"
            out = config.worship_dir / "openlp_song_index_elkton.md"
            rebuild_song_index(db, out)

        if do_lb:
            db = config.lb_root / "songs" / "songs.sqlite"
            out = config.worship_dir / "openlp_song_index_lb.md"
            rebuild_song_index(db, out)

        print()

    print("Done.")


if __name__ == "__main__":
    main()
