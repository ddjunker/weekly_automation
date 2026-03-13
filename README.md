# Weekly Automation

The **Weekly Automation** project is a cross-platform system that gathers, processes, and disseminates weekly liturgy and worship materials for the United Parish of Elkton and LB locations. It centralizes all scripts, templates, and documentation into a single, version-controlled workflow.

This repository contains:

- Python scripts that retrieve weekly worship elements  
- Text cleaning modules for OpenLP, Bible texts, and source websites  
- Obsidian documentation for planning and message preparation  
- Template files for bulletins, speaker notes, liturgist notes, and OpenLP  
- Config files that define paths, environment variables, and runtime behavior

The project is designed to run on Windows, Linux, and mixed-device environments used by the church’s technical infrastructure.

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

## Folder Structure (Full details in `STRUCTURE.md`)
```
docs/           → Master files, planning notes, documentation  
scripts/        → Python/PowerShell scripts for gathering, cleaning, and dissemination  
scripts/utils/  → Shared utility modules (I/O, XML, OpenLP helpers, text cleaning)  
templates/      → Bulletin templates, markdown skeletons, service structures  
data/           → Paths, logs, local OpenLP databases, temporary artifacts
```

## Setup

### Requirements
- Python 3.13  
- PowerShell  
- Git  
- LibreOffice  
- Obsidian  

### Installation
Clone the repository:

```bash
git clone git@github.com:ddjunker/weekly_automation.git
cd weekly_automation
```

Create a local `openlp_env.py` file in `scripts/utils/`:

```python
# Example format
OPENLP_PATH_ELK = r"PATH_TO_ELK_OPENLP"
OPENLP_PATH_LB = r"PATH_TO_LB_OPENLP"
```

*(This file is intentionally ignored by Git.)*

## Running the Scripts

Example workflow:

```powershell
python scripts/text_gather.py --master "docs/Master 2025-11-23.md"
python scripts/music_gather.py --master "docs/Master 2025-11-23.md"
python scripts/welcome.py
python -m scripts.publish "Master 2025-11-23.md"
```

## Rollback Point

Current milestone rollback tag:

- `beta-ready`

Useful commands:

```powershell
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
