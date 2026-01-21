# scripts/publish.py

from __future__ import annotations

import logging
import re
import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Tuple

from scripts.utils.config import config
from scripts.utils.file_io import read_text, write_text, safe_mkdir
from scripts.utils.logging_utils import init_logging
from scripts.utils.placeholder import extract_block
from scripts.utils.text_clean import clean_text, clean_markdown


# -----------------------------------------------------------------------------
# Configuration / naming
# -----------------------------------------------------------------------------

PASTOR_TEMPLATE_BY_COMMUNION = {
    "0": "Pastor.md",
    "1": "PastorC1.md",
    "2": "PastorC2.md",
    "3": "PastorC3.md",
}

DEFAULT_LITURGIST_TEMPLATE = "Liturgist.md"


@dataclass(frozen=True)
class RenderResult:
    output_path: Path
    warnings: Tuple[str, ...]


# -----------------------------------------------------------------------------
# Allow flexible input file reading
# -----------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Publish weekly worship materials from a master markdown file"
    )
    parser.add_argument(
        "master",
        help="Master markdown filename (relative to Worship directory)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Fail on unknown placeholders instead of warning",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    return parser.parse_args()



# -----------------------------------------------------------------------------
# Master placeholder parsing
# -----------------------------------------------------------------------------

PLACEHOLDER_TOKEN_RE = re.compile(r"\{([a-zA-Z0-9_]+)\}")


def extract_placeholders_used(template_text: str) -> Tuple[str, ...]:
    """Return unique placeholder names found in a template, in appearance order."""
    seen = set()
    ordered = []
    for m in PLACEHOLDER_TOKEN_RE.finditer(template_text):
        name = m.group(1)
        if name not in seen:
            ordered.append(name)
            seen.add(name)
    return tuple(ordered)


def build_master_lookup(master_md: str, placeholder_names: Iterable[str]) -> Dict[str, str]:
    """
    For each placeholder name, extract its block from Master.md.

    Notes:
      - Master.md stores blocks like:
            {placeholder}
            value...
      - extract_block returns the content after the placeholder until the next {x} or EOF.
    """
    lookup: Dict[str, str] = {}
    for name in placeholder_names:
        lookup[name] = extract_block(master_md, name)
    return lookup


# -----------------------------------------------------------------------------
# Template rendering rules
# -----------------------------------------------------------------------------

def _apply_omit_rule(rendered: str, placeholder_name: str, value: str) -> str:
    """
    If the replacement value is exactly 'omit' (case-insensitive),
    remove the *entire line* that contains the placeholder token.
    """
    if value.strip().lower() != "omit":
        return rendered

    # Remove lines containing {placeholder_name}
    pattern = re.compile(rf"^.*\{{{re.escape(placeholder_name)}\}}.*\n?", re.MULTILINE)
    return pattern.sub("", rendered)


def _apply_delete_rule(value: str) -> str:
    """
    If Master contains 'delete', replace with empty string.
    This mirrors your Master.md conventions.
    """
    return "" if value.strip().lower() == "delete" else value


def _resolve_aliases(name: str) -> Tuple[str, ...]:
    """
    Support known alias mismatches without forcing template edits.
    (We can remove aliases later once templates are normalized.)

    Current known mismatch:
      - templates sometimes use announce_details_* but mapping uses announce_detail_*
    """
    aliases = [name]

    # announce_details_* -> announce_detail_*
    if name.startswith("announce_details_"):
        aliases.append(name.replace("announce_details_", "announce_detail_", 1))

    return tuple(aliases)


def render_markdown_template(
    template_text: str,
    master_md: str,
    *,
    strict: bool = False,
) -> Tuple[str, Tuple[str, ...]]:
    """
    Render a markdown template by replacing placeholders with blocks extracted from Master.md.

    strict=False:
      - unknown placeholders remain untouched, and we emit a warning.
    strict=True:
      - unknown placeholders raise ValueError.

    Also applies:
      - omit rule: if replacement == 'omit', delete entire line containing placeholder
      - delete rule: if replacement == 'delete', replace with empty string
    """
    warnings = []

    placeholders = extract_placeholders_used(template_text)

    # Build lookup for placeholders used
    lookup = build_master_lookup(master_md, placeholders)

    rendered = template_text

    for name in placeholders:
        # Try aliases (for known drift between template and mapping/master)
        candidates = _resolve_aliases(name)

        value = ""
        found = False
        for candidate in candidates:
            candidate_value = lookup.get(candidate, "")
            if candidate_value:
                value = candidate_value
                found = True
                break

        if not found:
            msg = f"Placeholder {{{name}}} not found in Master.md (left unchanged)."
            if strict:
                raise ValueError(msg)
            warnings.append(msg)
            continue

        value = _apply_delete_rule(value)

        # Apply omit before actual replacement (needs the token still present)
        rendered = _apply_omit_rule(rendered, name, value)

        # Replace *all* occurrences of the placeholder token with the value
        # (templates sometimes repeat IDs/titles)
        rendered = rendered.replace(f"{{{name}}}", value)

    # Final cleanup for markdown stability (don’t reflow paragraphs)
    rendered = clean_markdown(rendered)
    return rendered, tuple(warnings)


# -----------------------------------------------------------------------------
# Output paths
# -----------------------------------------------------------------------------

def _get_date_slug(master_md: str) -> str:
    """
    Use {cal_date} if present; otherwise fallback.
    We avoid being clever here—just sanitize for filename.
    """
    raw = extract_block(master_md, "cal_date").strip()
    raw = raw or "unknown-date"
    raw = clean_text(raw)
    # filename-safe-ish
    raw = re.sub(r"[^\w\-]+", "-", raw)
    raw = re.sub(r"-{2,}", "-", raw).strip("-")
    return raw.lower()


def _speaker_output_path(worship_dir: Path, date_slug: str, template_name: str) -> Path:
    stem = Path(template_name).stem
    return worship_dir / f"{stem}--{date_slug}.md"


def _liturgist_output_path(worship_dir: Path, date_slug: str) -> Path:
    return worship_dir / f"Liturgist--{date_slug}.md"



# -----------------------------------------------------------------------------
# Public API: build files
# -----------------------------------------------------------------------------

def build_markdown_outputs(
    *,
    master_path: Path | None = None,
    templates_dir: Path | None = None,
    strict: bool = False,
    verbose: bool = False,
) -> Tuple[RenderResult, RenderResult]:
    """
    Generate:
      1) one Pastor/PastorC* speaker markdown file
      2) Liturgist markdown file

    Returns (speaker_result, liturgist_result)
    """
    init_logging(verbose=verbose)

    worship_dir = config.worship_dir

    if master_path is None:
        raise ValueError("master_path must be provided")
    if templates_dir is None:
        templates_dir = worship_dir / "templates"

    if not master_path.exists():
        raise FileNotFoundError(f"Master.md not found at: {master_path}")

    master_md = read_text(master_path)

    communion = extract_block(master_md, "communion").strip() or "0"
    template_name = PASTOR_TEMPLATE_BY_COMMUNION.get(communion, "Pastor.md")

    speaker_template_path = templates_dir / template_name
    liturgist_template_path = templates_dir / DEFAULT_LITURGIST_TEMPLATE

    if not speaker_template_path.exists():
        raise FileNotFoundError(f"Speaker template not found: {speaker_template_path}")
    if not liturgist_template_path.exists():
        raise FileNotFoundError(f"Liturgist template not found: {liturgist_template_path}")

    date_slug = _get_date_slug(master_md)

    logging.info("Worship dir: %s", worship_dir)
    logging.info("Master path: %s", master_path)
    logging.info("Speaker template path: %s", speaker_template_path)
    logging.info("Communion: %s -> template %s", communion, template_name)

    # Render speaker
    speaker_template = read_text(speaker_template_path)
    speaker_rendered, speaker_warnings = render_markdown_template(
        speaker_template,
        master_md,
        strict=strict,
    )
    speaker_out = _speaker_output_path(worship_dir, date_slug, template_name)
    write_text(speaker_out, speaker_rendered)

    # Render liturgist
    liturgist_template = read_text(liturgist_template_path)
    liturgist_rendered, liturgist_warnings = render_markdown_template(
        liturgist_template,
        master_md,
        strict=strict,
    )
    liturgist_out = _liturgist_output_path(worship_dir, date_slug)
    write_text(liturgist_out, liturgist_rendered)

    # Log warnings
    for w in speaker_warnings:
        logging.warning("[Speaker] %s", w)
    for w in liturgist_warnings:
        logging.warning("[Liturgist] %s", w)

    return (
        RenderResult(speaker_out, speaker_warnings),
        RenderResult(liturgist_out, liturgist_warnings),
    )


# -----------------------------------------------------------------------------
# CLI entry point
# -----------------------------------------------------------------------------

if __name__ == "__main__":
    args = parse_args()

    worship_dir = config.worship_dir
    master_path = worship_dir / args.master

    speaker, liturgist = build_markdown_outputs(
        master_path=master_path,
        strict=args.strict,
        verbose=args.verbose,
    )

    print(f"Speaker file:   {speaker.output_path.name}")
    print(f"Liturgist file: {liturgist.output_path.name}")

    if speaker.warnings:
        print("\nSpeaker warnings:")
        for w in speaker.warnings:
            print(" -", w)

    if liturgist.warnings:
        print("\nLiturgist warnings:")
        for w in liturgist.warnings:
            print(" -", w)
