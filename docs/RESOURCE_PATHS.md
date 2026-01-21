# Resource Paths Reference

This document lists the expected subdirectory layout under the OpenLP roots and other resource locations used by the scripts.

## Overview
- `elkton_openlp_root` and `lb_openlp_root` should point to the top of each OpenLP installation (the folder that contains `songs/`, `bibles/`, `custom/`, etc.).
- The scripts discover the song DB at `<root>/songs/songs.sqlite` and Bible DB files under `<root>/bibles/*.sqlite`.

## Expected structure under each OpenLP root
- `<openlp_root>/songs/` — contains `songs.sqlite` (song database)
- `<openlp_root>/bibles/` — one or more `.sqlite` files for Bibles (e.g. `Good News Translation.sqlite`, `New Revised Standard Version.sqlite`)
- `<openlp_root>/custom/` — `custom.sqlite` for custom slides
- `<openlp_root>/exported_slides/` — exported slides (used by some workflows)
- `<openlp_root>/OpenLP Services/` — .osz service packages
- `<openlp_root>/themes/` — theme folders (Advent-Christmas, Default, Transparent, etc.)

## Other project-level resources
- `worship_dir` — Markdown bulletins and templates (Windows example: `C:\Users\Public\Documents\Area42\Worship`)
- `ccli_dir` — CCLI song/metadata folder
- `browser_profile` — (optional) Playwright persistent profile stored at repository root `browser_profile/`.

### `worship_dir` subfolders used by dissemination scripts

- `templates/` — speaker and liturgist templates. Keep master templates here so the dissemination script can copy/populate them.
- `resources/` — static assets used in speaker files. The dissemination script expects these to be available under `worship_dir/resources` when assembling output.

## Notes for developers
- Prefer configuring only the OpenLP roots in `scripts/weekly_config.ini`; the scripts will resolve DB locations relative to those roots.
- Use `scripts/weekly_config.ini.sample` as a starting point for per-machine `scripts/weekly_config.ini`.
- If tools require repo-relative paths, consider creating junctions/symlinks on Windows/Linux instead of modifying code.

