"""
Shared command-line plumbing.

Every script takes the same master argument and sets up logging the same way;
each had its own copy, including a legacy `--master` flag kept alongside the
positional one. Commit 99fdfa9 dropped that flag from publish; this finishes the
job for the rest.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from scripts.utils.config import resolve_master_path
from scripts.utils.logging_utils import init_logging


def add_master_arg(parser: argparse.ArgumentParser) -> None:
    """The master markdown argument, spelled the same way everywhere."""
    parser.add_argument(
        "master",
        nargs="?",
        default="",
        help="Master markdown filename (relative to the Worship directory) or path",
    )


def add_verbose_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--verbose", action="store_true", help="Enable debug logging")


def resolve_master_or_exit(args: argparse.Namespace, *, script: str) -> Path:
    """
    Turn the parsed argument into an existing master path, or exit with a message
    naming the script, so the error reads the same from every entry point.
    """
    master_arg = str(getattr(args, "master", "") or "").strip()
    if not master_arg:
        raise SystemExit(
            f'Master markdown path is required (e.g. {script} "M10-04.md")'
        )

    master_path = resolve_master_path(master_arg)
    if not master_path.exists():
        raise SystemExit(f"Master file not found: {master_path}")
    return master_path


def start_logging(args: argparse.Namespace) -> None:
    init_logging(verbose=bool(getattr(args, "verbose", False)))
