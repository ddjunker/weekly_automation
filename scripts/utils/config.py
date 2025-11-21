"""
utils/config.py
Central configuration loader for Weekly Automation scripts.
"""

from __future__ import annotations

import configparser
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

        paths = self._parser["paths"]

        # Exported values (all converted to Path objects)
        self.worship_dir = Path(paths["worship_dir"])
        self.ccli_dir = Path(paths["ccli_dir"])

        self.elkton_root = Path(paths["elkton_openlp_root"])
        self.lb_root = Path(paths["lb_openlp_root"])

        self.elkton_songs_db = Path(paths["elkton_songs_db"])
        self.lb_songs_db = Path(paths["lb_songs_db"])

        self.elkton_bible_db = Path(paths["elkton_bible_db"])
        self.lb_bible_db = Path(paths["lb_bible_db"])


# Instantiate global singleton
config = Config()
