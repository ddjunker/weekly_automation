# Weekly Automation — v1.1.2

The **Weekly Automation** project is a cross-platform system that gathers, processes, and disseminates weekly liturgy and worship materials for the United Parish of Elkton and LB locations. It centralizes all scripts, templates, and documentation into a single, version-controlled workflow.

This repository contains:

- Python scripts that retrieve weekly worship elements  
- Text cleaning modules for OpenLP, Bible texts, and source websites  
- Obsidian documentation for planning and message preparation  
- Template files for bulletins, speaker notes, liturgist notes, and OpenLP  
- Config files that define paths, environment variables, and runtime behavior

The project is designed to run on Windows, Linux, and mixed-device environments used by the church's technical infrastructure.

## Features

### 🔍 Gathering
- Fetches Scripture texts, Calls to Worship, Benedictions, and other weekly elements  
- Pulls XML texts from existing OpenLP databases  
- Optional module to retrieve song metadata or titles from two separate OpenLP installations

### 🧽 Cleaning & Normalization
- Unicode scrubbing  
- Markdown normalization  
- Whitespace and formatting correction  
- HTML → clean-text transformations  
- Slide/text harmonization for OpenLP

### 📄 Dissemination
- Generates markdown output for:
  - Speaker notes  
  - Liturgist notes  
- Populates LibreOffice `.odt` bulletin templates  
- Prepares OpenLP elements (planned future integration)

## Folder Structure
```
docs/           → Planning notes and path documentation
scripts/        → The scripts themselves
scripts/utils/  → Shared modules: config, file I/O, placeholders, text cleaning,
                  OpenLP databases (openlp.py), service files (osz.py),
                  hymnal references (hymnal.py), run reports (report.py)
tests/          → pytest suite; tests/data holds trimmed service files
```
Templates, master files and the OpenLP databases live outside the repo, under the
paths set in `scripts/weekly_config.ini`.

## Tests
```bash
pip install -r requirements-dev.txt
python -m pytest
```
The suite never touches the live OneDrive tree: service files are rebuilt from the
trimmed fixtures in `tests/data`, and the databases are created in a temp folder.
Regenerate the fixtures after a template change with
`python3 tests/tools/make_osz_fixture.py`.

## Setup

### Requirements
- Python 3.12+
- Git
- LibreOffice
- Obsidian
- GIMP 3 (for `welcome.py` — see Linux notes below for installation)

### Installation

Clone the repository (public — no authentication required):

```bash
git clone https://github.com/ddjunker/weekly_automation.git
cd weekly_automation
```

Create a virtual environment. Use `--system-site-packages` so GIMP's Python bindings (`gi`) are accessible:

```bash
python3 -m venv --system-site-packages .venv
source .venv/bin/activate          # Linux/macOS
.venv\Scripts\activate             # Windows
```

Install Python dependencies:

```bash
pip install -r requirements.txt
playwright install firefox
```

### Configuration

`weekly_config.ini` is committed to the repository and includes platform-specific sections. On Linux, values in `[paths_linux]` override `[paths]` automatically — no manual edits needed unless your paths differ from the defaults.

### Linux Deployment

**GIMP 3**

Ubuntu's default apt repository includes GIMP 2.10, which is not compatible with this project. Install GIMP 3 via the dedicated PPA:

```bash
sudo add-apt-repository ppa:ubuntuhandbook1/gimp-3
sudo apt update
sudo apt install gimp
```

Verify: `gimp --version` should report 3.x. No `gimp_bin_dir` setting is needed — the script finds GIMP automatically via PATH.

**Clipboard support for `text_gather.py -w`**

The `-w` flag reads from the system clipboard. Install the appropriate tool:

```bash
# X11 sessions
sudo apt install xclip

# Wayland sessions
sudo apt install wl-clipboard
```

**Virtual environment activation**

Add to `~/.bashrc` or `~/.profile` to activate automatically on login:

```bash
source ~/weekly_automation/.venv/bin/activate
```

## Running the Scripts

Example workflow:

```bash
# 1. Pre-publication checks (read-only, writes report files)
python3 -m scripts.prepub "Master 2025-11-23.md"

# 2. Gather text and music into the master file
python3 -m scripts.text_gather "Master 2025-11-23.md"
python3 -m scripts.music_gather "Master 2025-11-23.md"

# 3. Export welcome slides
python3 -m scripts.welcome "Master 2025-11-23.md"

# 4. Publish all outputs
python3 -m scripts.publish "Master 2025-11-23.md"
```

### prepub.py — pre-publication checks

Validates all resources against local databases without writing to the master file. Run this first to catch missing or ambiguous entries before gathering.

| Flag | Checks |
|------|--------|
| `-t` | Scripture texts, CtW and AoF custom slides |
| `-m` | Songs (hymnal DB and CCLI lyrics) |

Omitting both flags runs all checks. Writes a `_text_check.md` and/or `_music_check.md` report next to the master file.

### Call to Worship — per-church exact matching

Some Psalms have multiple CtW slide variants in OpenLP (e.g., `CtW Psalm 23:1-6` and `CtW Psalm 23:1-6 UMH 128`). To resolve ambiguity:

1. Add two blank placeholders to your Obsidian master template (once):
   ```
   {ctw_ref_elkton}
   {ctw_ref_lb}
   ```
2. Run `prepub.py -t`. If the report shows `MULTIPLE` for a church, copy the exact slide title from the Detail column into the appropriate placeholder.
3. Run the real gather (`-c`) and publish — both use exact matching on those values. If a per-church placeholder is blank, CtW is skipped for that church.

`{ctw_ref}` (the lectionary reference) is still used by the dry run to enumerate candidates. The per-church placeholders hold the full OpenLP slide title selected for that week.

### text_gather.py section flags

By default `text_gather.py` runs all sections. Use one or more of the following flags to update only specific placeholders:

| Flag | Updates |
|------|---------|
| `-s` | Scripture texts |
| `-c` | Call to Worship (CtW) slides |
| `-a` | Affirmation of Faith (AoF) slides |
| `-w` | Offertory and benediction (launches Firefox) |

Flags may be combined. Omitting all flags runs everything.

```bash
# Scripture only
python3 -m scripts.text_gather "Master 2025-11-23.md" -s

# CtW and AoF only (no browser)
python3 -m scripts.text_gather "Master 2025-11-23.md" -c -a

# Offertory and benediction only
python3 -m scripts.text_gather "Master 2025-11-23.md" -w
```

### publish.py output switches

By default `publish.py` generates all three output types. Use one of the following flags to run only a specific segment:

| Flag | Long form | Output |
|------|-----------|--------|
| `-m` | `--markdown` | Speaker and liturgist markdown files only |
| `-w` | `--writer` | LibreOffice Writer `.odt` files only |
| `-o` | `--openlp` | OpenLP `.osz` service files only |

These flags are mutually exclusive. Omitting all three runs everything.

`--dry-run` reports what the OpenLP injection would do without writing any service
file, and `--strict` makes the run exit non-zero when anything was missed. Every run
ends with a summary naming what matched, what was deliberately skipped and what
failed; output files are still written when something fails, so they can be fixed by
hand in OpenLP, and the summary marks them as containing template content.

```
=== OpenLP summary ===
elkton   OK              11 ok / 2 skip / 0 warn / 0 error   -> Service--2026-10-04.osz
   SKIP   aof                  no aof_ref set in the master
   SKIP   song_opening         id 'MercyMe' is not a hymnal reference
```

```bash
# Markdown only
python3 -m scripts.publish "Master 2025-11-23.md" -m

# Writer files only
python3 -m scripts.publish "Master 2025-11-23.md" -w

# OpenLP only
python3 -m scripts.publish "Master 2025-11-23.md" -o
```

### Occasional scripts

| Script | Purpose |
|--------|---------|
| `rebuild_indexes.py` | Rewrite `openlp_custom_index_<church>.md` from the custom slide libraries. `--slides` / `--elkton` / `--lb` narrow it. Song index generation is not implemented and writes nothing. |
| `dedupe_custom_slides.py` | Remove duplicate custom slides (same title and credits) that older runs made OpenLP clone. Dry run by default; `--apply` deletes, after a timestamped backup. |

## Backups

`text_gather.py` and `music_gather.py` rewrite the master in place, so each run first
copies it to `.backups/` beside the master, keeping the five most recent. `publish.py`
backs up a church's `custom.sqlite` once per run before it writes this week's offertory
prayer and benediction into the library. All output files are written through a temp
file and renamed into place, so an interrupted run cannot leave a half-written
bulletin or service file.

## Rollback Point

Current milestone rollback tag:

- `beta-ready`

Useful commands:

```bash
# Inspect the tagged snapshot
git checkout beta-ready

# Return to main
git checkout main

# Reset local main to the rollback point (destructive to local commits after the tag)
git checkout main
git reset --hard beta-ready
```

## License
Internal project for United Parish of Elkton. Not intended for redistribution.

## Maintainer
**Daren Junker**
