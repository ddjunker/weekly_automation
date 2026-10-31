"""
utils/config.py
Central configuration loader for Weekly Automation scripts.
"""

from __future__ import annotations

import configparser
import os
import platform
from pathlib import Path


class Config:
    """Singleton configuration reader."""

    def __init__(self):
        # Path to weekly_config.ini is always one directory above utils/
        self.config_path = Path(__file__).parent.parent / "weekly_config.ini"

        if not self.config_path.exists():
            raise FileNotFoundError(
                f"Configuration file not found: {self.config_path}"
            )

        self._parser = configparser.ConfigParser()
        self._parser.read(self.config_path, encoding="utf-8")

        # Load the generic [paths] section, then overlay an OS-specific
        # section if present (e.g. [paths_linux], [paths_windows]).
        paths: dict = {}
        if "paths" in self._parser:
            paths.update(self._parser["paths"])

        sys_name = platform.system().lower()
        platform_section = f"paths_{sys_name}"
        if platform_section in self._parser:
            paths.update(self._parser[platform_section])

        # Expand env vars and ~, returning a Path — or None when the key is blank.
        # It used to return Path(""), which is Path("."): truthy, and .exists() is
        # True, so a `if not some_dir` guard could never fire and a missing
        # bulletin_dir would have written .odt files into the current directory.
        def _p(key: str) -> Path | None:
            raw = (paths.get(key) or "").strip()
            if not raw:
                return None
            return Path(os.path.expanduser(os.path.expandvars(raw)))

        # Exported values (all converted to Path objects)

        self.worship_dir = _p("worship_dir")
        self.ccli_dir = _p("ccli_dir")
        self.bulletin_dir = _p("bulletin_dir")

        self.elkton_root = _p("elkton_openlp_root")
        self.lb_root = _p("lb_openlp_root")

        # Song and Bible databases are always discovered under the OpenLP roots
        # (see utils/openlp.py), so the *_songs_db / *_bible_db ini keys are not read.

        # Welcome slide / GIMP integration
        self.projection_pics_dir = _p("projection_pics_dir")
        self.gimp_bin_dir = _p("gimp_bin_dir")

        # ----------------------------------------------------------
        # NEW FIELD: browser profile directory for Playwright
        # ----------------------------------------------------------
        # A persistent browser profile is needed so Playwright retains:
        # - clipboard permissions
        # - cookies / sessions (if ever needed)
        # - consistent behavior between script runs
        #
        # Stored in the project root (same as weekly_automation/)
        # Automatically created by Playwright if it doesn't exist.
        self.browser_profile = Path(__file__).resolve().parents[2] / "browser_profile"


# Instantiate global singleton
config = Config()


def require_dir(path: Path | None, key: str) -> Path:
    """
    Return an existing directory or explain which weekly_config.ini key to set.
    One message, instead of each caller inventing its own check.
    """
    if path is None:
        raise FileNotFoundError(f"weekly_config.ini has no value for '{key}'")
    expanded = path.expanduser()
    if not expanded.exists():
        raise FileNotFoundError(f"Directory for '{key}' not found: {expanded}")
    return expanded


def resolve_master_path(master_arg: str) -> Path:
    """
    Resolve a Master markdown path:
      1. Absolute path → returned as-is.
      2. Relative path → resolved under config.worship_dir.
    worship_dir must exist (strict).
    """
    mp = Path(master_arg)
    if mp.is_absolute():
        return mp
    worship = config.worship_dir.expanduser().resolve()
    if not worship.exists():
        raise FileNotFoundError(f"Worship directory not found: {worship}")
    return worship / master_arg
