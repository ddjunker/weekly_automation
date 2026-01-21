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

        # Helper to expand env vars and ~, returning a Path. Missing keys
        # yield an empty Path() which is falsy for .exists() checks.
        def _p(key: str) -> Path:
            raw = paths.get(key, "")
            if not raw:
                return Path("")
            expanded = os.path.expandvars(raw)
            expanded = os.path.expanduser(expanded)
            return Path(expanded)

        # Exported values (all converted to Path objects)
        self.worship_dir = _p("worship_dir")
        self.ccli_dir = _p("ccli_dir")

        self.elkton_root = _p("elkton_openlp_root")
        self.lb_root = _p("lb_openlp_root")

        # These may be left blank in favor of root-relative discovery.
        self.elkton_songs_db = _p("elkton_songs_db")
        self.lb_songs_db = _p("lb_songs_db")

        self.elkton_bible_db = _p("elkton_bible_db")
        self.lb_bible_db = _p("lb_bible_db")

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
