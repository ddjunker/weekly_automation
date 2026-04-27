#!/usr/bin/env python3
"""
welcome.py
GIMP 3.x headless updater: set cal_date + cal_church text layers and export PNG.

Uses existing weekly_automation infrastructure for config, placeholder extraction,
and file handling.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, cast

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

try:
    from scripts.utils.config import config
    from scripts.utils.file_io import read_text
    from scripts.utils.logging_utils import init_logging
except (ImportError, ModuleNotFoundError):
    # Fallback for embedded runtimes (e.g., GIMP Python) where importing
    # scripts.utils package-level __init__ may fail due to optional deps
    # (e.g., sqlite3 DLL unavailable, yaml/regex not installed).
    utils_dir = Path(__file__).resolve().parent / "utils"
    sys.path.insert(0, str(utils_dir))
    from config import config  # type: ignore[import-not-found]
    from file_io import read_text  # type: ignore[import-not-found]
    from logging_utils import init_logging  # type: ignore[import-not-found]


def _iter_layers_recursive(layer):
    """Yield a layer; if it's a group, yield its children recursively."""
    yield layer
    # Group layers expose children; this is the most robust way to detect.
    if hasattr(layer, "get_children"):
        for child in layer.get_children():
            yield from _iter_layers_recursive(child)


def _find_layer_by_name(image: Any, name: str) -> Any | None:
    for top in image.get_layers():
        for layer in _iter_layers_recursive(top):
            if layer.get_name() == name:
                return layer
    return None


def _setup_gimp_runtime() -> tuple[Any, Any]:
    """
    Configure runtime search paths for GIMP/gi on Windows and import modules lazily.
    Returns (Gimp, Gio).
    """
    gimp_bin = getattr(config, "gimp_bin_dir", Path(""))
    if gimp_bin:
        gimp_bin_str = str(gimp_bin)
        os.environ["PATH"] = f"{gimp_bin_str}{os.pathsep}" + os.environ.get("PATH", "")
        if os.name == "nt" and hasattr(os, "add_dll_directory"):
            try:
                os.add_dll_directory(gimp_bin_str)
            except OSError:
                logging.debug("Could not add GIMP bin to DLL search path: %s", gimp_bin_str)

    import gi  # type: ignore[import-not-found]

    gi.require_version("Gimp", "3.0")
    gi.require_version("Gio", "2.0")
    from gi.repository import Gimp, Gio  # type: ignore[import-not-found]

    return Gimp, Gio


def _find_gimp_console_exe() -> Path | None:
    import shutil
    gimp_bin = getattr(config, "gimp_bin_dir", Path(""))
    names = ["gimp-console-3.0", "gimp-console", "gimp-3.0", "gimp"]
    if gimp_bin:
        if os.name == "nt":
            candidates = [gimp_bin / f"{n}.exe" for n in names]
        else:
            candidates = [gimp_bin / n for n in names]
        for candidate in candidates:
            if candidate.exists():
                return candidate
    # Fall back to PATH search
    for name in names:
        found = shutil.which(name)
        if found:
            return Path(found)
    return None


def _ensure_gimp_runtime_or_reinvoke(argv: list[str]) -> None:
    """
    Re-invoke this script through GIMP's python-fu batch interpreter
    unless we are already running inside it (WA_GIMP_BATCH=1).
    """
    if os.environ.get("WA_GIMP_BATCH") == "1":
        return

    gimp_console = _find_gimp_console_exe()
    if not gimp_console:
        raise RuntimeError(
            "No GIMP executable found. Set gimp_bin_dir in weekly_config.ini "
            "or ensure GIMP is on your PATH."
        )

    repo_root = Path(__file__).resolve().parents[1]
    py_batch = (
        "import sys; "
        f"sys.path.insert(0, {str(repo_root)!r}); "
        "import scripts.welcome as w; "
        f"w.main({argv!r})"
    )

    cmd = [
        str(gimp_console),
        "-i",
        "--batch-interpreter=python-fu-eval",
        "--batch",
        py_batch,
        "--quit",
    ]

    env = os.environ.copy()
    env["WA_GIMP_BATCH"] = "1"
    subprocess.run(cmd, check=True, env=env)
    raise SystemExit(0)


def _resolve_path(value: str, base_dir: Path | None = None) -> Path:
    p = Path(value).expanduser()
    if p.is_absolute():
        return p
    if base_dir is not None:
        return (base_dir / p).resolve()
    return p.resolve()


def _resolve_template_path(template_arg: str, projection_dir: Path) -> Path:
    """
    Resolve template .xcf path.
    - If template_arg is provided, resolve it directly.
    - If omitted, auto-detect a common welcome template in projection_dir.
    """
    template_arg = (template_arg or "").strip()
    if template_arg:
        return _resolve_path(template_arg, projection_dir)

    candidates = (
        "Worship Welcome.xcf",
        "Worship Welcome LB.xcf",
        "welcome_template.xcf",
    )
    for name in candidates:
        candidate = projection_dir / name
        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        "No default welcome template found in projection_pics_dir. "
        f"Checked: {', '.join(candidates)}"
    )


def _resolve_template_paths(template_arg: str, projection_dir: Path) -> list[Path]:
    """
    Resolve one or more template paths.
    - If template_arg is provided, returns a single-item list.
    - If omitted, returns all known church welcome templates that exist.
    """
    template_arg = (template_arg or "").strip()
    if template_arg:
        return [_resolve_template_path(template_arg, projection_dir)]

    candidates = (
        projection_dir / "Worship Welcome.xcf",
        projection_dir / "Worship Welcome LB.xcf",
    )
    found = [p for p in candidates if p.exists()]
    if found:
        return found

    # Fallback to legacy/default single-template discovery.
    return [_resolve_template_path("", projection_dir)]


def _resolve_master_path(master_arg: str) -> Path:
    """
    Resolve master path similarly to text_gather.py:
      - absolute path as-is
      - otherwise resolve relative to config.worship_dir
    """
    candidate = Path(master_arg).expanduser()
    if candidate.is_absolute():
        return candidate

    worship = config.worship_dir.expanduser().resolve()
    if not worship.exists():
        raise FileNotFoundError(f"Worship directory not found: {worship}")

    return (worship / candidate).resolve()



def _extract_block(md: str, key: str) -> str:
    """Extract value after =={key}== (or bare {key}). Uses only stdlib re."""
    pattern = rf"\{{{re.escape(key)}\}}=*\s*(.*?)(?=\n\s*=*\{{|\Z)"
    m = re.search(pattern, md, flags=re.S)
    if not m:
        return ""
    text = re.sub(r"%%[\s\S]*?%%", "", m.group(1))  # strip Obsidian comments
    return text.strip()


def _load_welcome_fields(master_path: Path) -> tuple[str, str]:
    if not master_path.exists():
        raise FileNotFoundError(f"Master markdown not found: {master_path}")
    md = read_text(master_path)
    cal_date = _extract_block(md, "cal_date").strip()
    cal_church = _extract_block(md, "cal_church").strip()
    if not cal_date:
        raise ValueError("{cal_date} is missing in Master markdown")
    if not cal_church:
        raise ValueError("{cal_church} is missing in Master markdown")
    return cal_date, cal_church



def _set_text_layer_content(layer: Any, value: str, layer_name: str) -> None:
    text_value = (value or "").strip()

    def _escape_pango_text(text: str) -> str:
        escaped = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        return escaped.replace("\n", "&#10;")

    def _inject_text_into_markup(markup: str, text: str) -> str:
        escaped_text = _escape_pango_text(text)
        token_re = re.compile(r">([^<]*)<")

        replaced = False

        def _repl(match: re.Match[str]) -> str:
            nonlocal replaced
            content = match.group(1)
            if content.strip() and not replaced:
                replaced = True
                return f">{escaped_text}<"
            if content.strip():
                return "><"
            return match.group(0)

        result = token_re.sub(_repl, markup)
        if replaced:
            return result

        close_idx = result.rfind("</")
        if close_idx >= 0:
            return result[:close_idx] + escaped_text + result[close_idx:]
        return escaped_text

    def _build_markup_from_template(text: str) -> str:
        plain_text = text.replace("\r\n", "\n").replace("\r", "\n")
        escaped_text = _escape_pango_text(plain_text)
        if not hasattr(layer, "get_markup"):
            return escaped_text

        try:
            template_markup = str(layer.get_markup() or "")
        except Exception:
            return escaped_text

        if not template_markup:
            return escaped_text

        return _inject_text_into_markup(template_markup, plain_text)

    def _current_text() -> str:
        if hasattr(layer, "get_text"):
            try:
                return str(layer.get_text() or "")
            except Exception:
                return ""
        return ""

    def _current_markup() -> str:
        if hasattr(layer, "get_markup"):
            try:
                return str(layer.get_markup() or "")
            except Exception:
                return ""
        return ""

    set_ok = False
    if hasattr(layer, "set_markup"):
        markup_text = _build_markup_from_template(text_value)
        try:
            set_ok = bool(layer.set_markup(markup_text))
            if set_ok and _current_markup().strip():
                return
        except Exception:
            set_ok = False

    current = _current_text().strip()
    if text_value and current:
        return

    if hasattr(layer, "set_text"):
        try:
            set_ok = bool(layer.set_text(text_value))
            current = _current_text().strip()
            if text_value and current:
                return
        except Exception:
            pass

    raise RuntimeError(
        f"Failed to set text for layer '{layer_name}'. set_ok={set_ok}, "
        f"input={text_value!r}, current={current!r}, markup={_current_markup()!r}"
    )


def update_welcome_slide_xcf_to_png(
    template_xcf: str,
    output_png: str,
    cal_date_text: str,
    cal_church_text: str,
    layer_date_name: str = "cal_date",
    layer_church_name: str = "cal_church",
):
    template_xcf = str(template_xcf)
    output_png = str(output_png)

    Gimp, Gio = _setup_gimp_runtime()

    # Load template
    in_file = Gio.File.new_for_path(template_xcf)
    image = Gimp.file_load(Gimp.RunMode.NONINTERACTIVE, in_file)
    if image is None:
        raise RuntimeError(f"GIMP could not load template image: {template_xcf}")

    # Find layers
    layer_date = _find_layer_by_name(image, layer_date_name)
    layer_church = _find_layer_by_name(image, layer_church_name)

    missing = [n for n, l in [(layer_date_name, layer_date), (layer_church_name, layer_church)] if l is None]
    if missing:
        raise RuntimeError(f"Missing layer(s) in template: {missing}")

    # Safety check: they should be TextLayers
    if not isinstance(layer_date, Gimp.TextLayer):
        raise RuntimeError(f"Layer '{layer_date_name}' is not a TextLayer (it is {type(layer_date)})")
    if not isinstance(layer_church, Gimp.TextLayer):
        raise RuntimeError(f"Layer '{layer_church_name}' is not a TextLayer (it is {type(layer_church)})")

    layer_date = cast(Any, layer_date)
    layer_church = cast(Any, layer_church)

    # Set text (supports \n line breaks)
    _set_text_layer_content(layer_date, cal_date_text, layer_date_name)
    _set_text_layer_content(layer_church, cal_church_text, layer_church_name)

    # Export PNG (exporter chosen by file extension)
    out_file = Gio.File.new_for_path(output_png)
    Gimp.file_save(Gimp.RunMode.NONINTERACTIVE, image, out_file, None)

    # Clean up
    image.delete()


def main(argv=None):
    argv_list = list(sys.argv[1:] if argv is None else argv)
    _ensure_gimp_runtime_or_reinvoke(argv_list)

    parser = argparse.ArgumentParser(
        description="Update a GIMP welcome-slide template from Master.md and export PNG"
    )
    parser.add_argument(
        "master",
        nargs="?",
        default="",
        help="Master markdown path (absolute, repo-relative, or worship_dir-relative)",
    )
    parser.add_argument(
        "--master",
        dest="master_flag",
        default="",
        help="Master markdown path (same as positional master)",
    )
    parser.add_argument(
        "--template",
        default="",
        help="Template .xcf filename or absolute path (optional; auto-detects common names if omitted)",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Output .png filename or absolute path (single-template mode only)",
    )
    parser.add_argument("--layer-date", default="cal_date", help="Text layer name for date")
    parser.add_argument("--layer-church", default="cal_church", help="Text layer name for church")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")

    args = parser.parse_args(argv_list)
    init_logging(verbose=args.verbose)

    master_arg = (args.master_flag or args.master or "data/Master.md").strip()

    projection_dir = getattr(config, "projection_pics_dir", Path(""))
    if not projection_dir:
        raise ValueError("projection_pics_dir is not configured in weekly_config.ini")
    if not projection_dir.exists():
        raise FileNotFoundError(f"projection_pics_dir does not exist: {projection_dir}")

    master_path = _resolve_master_path(master_arg)
    cal_date_text, cal_church_text = _load_welcome_fields(master_path)
    logging.debug("Loaded cal_date=%r", cal_date_text)
    logging.debug("Loaded cal_church=%r", cal_church_text)

    template_paths = _resolve_template_paths(args.template, projection_dir)

    if args.output and len(template_paths) > 1:
        raise ValueError(
            "--output can only be used when rendering a single template. "
            "Use --template to target one file, or omit --output for multi-template mode."
        )

    for template_path in template_paths:
        if not template_path.exists():
            raise FileNotFoundError(f"Welcome slide template not found: {template_path}")

        if args.output:
            output_path = _resolve_path(args.output, projection_dir)
        else:
            output_path = template_path.with_suffix(".png")

        output_path.parent.mkdir(parents=True, exist_ok=True)

        update_welcome_slide_xcf_to_png(
            str(template_path),
            str(output_path),
            cal_date_text,
            cal_church_text,
            layer_date_name=args.layer_date,
            layer_church_name=args.layer_church,
        )
        logging.info("Exported welcome slide PNG: %s", output_path)


if __name__ == "__main__":
    main()