# Combined function for both markdown and writer outputs
from __future__ import annotations
import logging
import re
import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Tuple
from datetime import datetime
import json
import zipfile
import tempfile
from copy import deepcopy
from xml.etree import ElementTree as ET

# Add template mappings at the top so they are available to all functions
PASTOR_TEMPLATE_BY_COMMUNION = {
    "0": {
        "elkton_md": "Pastor.md",
        "lb_md": "PastorLB.md",
        "elkton_writer": "elk.ott",
        "lb_writer": "lb.ott",
        "elkton_openlp": "template_sunday.osz",
        "lb_openlp": "template_lb.osz",
    },
    "1": {
        "elkton_md": "PastorC1.md",
        "lb_md": "PastorLBC.md",
        "elkton_writer": "elkc1.ott",
        "lb_writer": "lbc.ott",
        "elkton_openlp": "template_communion1.osz",
        "lb_openlp": "template_lbcommunion.osz",
    },
    "2": {
        "elkton_md": "PastorC2.md",
        "lb_md": "PastorLBC.md",
        "elkton_writer": "elkc2.ott",
        "lb_writer": "lbc.ott",
        "elkton_openlp": "template_communion2.osz",
        "lb_openlp": "template_lbcommunion.osz",
    },
    "3": {
        "elkton_md": "PastorC3.md",
        "lb_md": "PastorLBC.md",
        "elkton_writer": "elkc3.ott",
        "lb_writer": "lbc.ott",
        "elkton_openlp": "template_communion3.osz",
        "lb_openlp": "template_lbcommunion.osz",
    },
}

DEFAULT_LITURGIST_TEMPLATE = "Liturgist.md"


# RenderResult dataclass for output tracking
@dataclass(frozen=True)
class RenderResult:
    output_path: Path
    warnings: Tuple[str, ...]


# -----------------------------------------------------------------------------
# OpenLP Template Copier
# -----------------------------------------------------------------------------

def _load_service_data_from_osz(osz_path: Path) -> tuple[list, dict[str, bytes], list[zipfile.ZipInfo]]:
    """Read service_data.osj and all archive members from an OpenLP .osz file."""
    with zipfile.ZipFile(osz_path, "r") as zin:
        infos = zin.infolist()
        payloads = {info.filename: zin.read(info.filename) for info in infos}

    if "service_data.osj" not in payloads:
        raise ValueError(f"service_data.osj not found in {osz_path}")

    service_data = json.loads(payloads["service_data.osj"].decode("utf-8"))
    return service_data, payloads, infos


def _write_service_data_to_osz(
    osz_path: Path,
    service_data: list,
    payloads: dict[str, bytes],
    infos: list[zipfile.ZipInfo],
) -> None:
    """Write updated service_data.osj back to an OpenLP .osz, preserving other members."""
    payloads["service_data.osj"] = json.dumps(service_data, ensure_ascii=False).encode("utf-8")

    tmp_path = osz_path.with_suffix(osz_path.suffix + ".tmp")
    with zipfile.ZipFile(tmp_path, "w") as zout:
        for info in infos:
            data = payloads.get(info.filename, b"")
            zout.writestr(info, data)

    tmp_path.replace(osz_path)


def _get_slide_text_for_prefix(church: str, prefix: str) -> tuple[str, str] | None:
    """
    Find a custom slide whose title starts with prefix and return (title, cleaned_text).
    """
    from scripts.utils.openlp import list_custom_slides, load_custom_slide
    from scripts.utils.text_clean import clean_text, xml_to_text

    prefix_clean = clean_text(prefix).lower()
    slides = list_custom_slides(church)

    matches = [s for s in slides if clean_text(s["title"]).lower().startswith(prefix_clean)]
    if not matches:
        return None

    # Prefer exact normalized title match if available.
    chosen = next(
        (s for s in matches if clean_text(s["title"]).lower() == prefix_clean),
        matches[0],
    )

    slide = load_custom_slide(church, chosen["uuid"])
    if not slide or not isinstance(slide.get("text"), str) or not slide["text"].strip():
        return None

    text = xml_to_text(slide["text"])
    text = re.sub(r"\[={3,}\]|\[[\-—]{3,}\]|\{y\}|\{/y\}", "", text).strip()
    if not text:
        return None

    return str(chosen["title"]), text


def _replace_custom_item_text(service_data: list, marker_title: str, new_title: str, new_text: str) -> bool:
    """
    Replace one custom service item (matched by header.title) with new custom-slide text.
    Returns True if an item was replaced.
    """
    marker = marker_title.strip().lower()

    for item in service_data:
        svc = item.get("serviceitem")
        if not isinstance(svc, dict):
            continue

        header = svc.get("header")
        if not isinstance(header, dict):
            continue

        if str(header.get("plugin", "")).lower() != "custom":
            continue

        title = str(header.get("title", "")).strip().lower()
        if title != marker:
            continue

        header["title"] = new_title
        header["footer"] = [f"{new_title} "]
        svc["data"] = [{
            "title": new_title[:30],
            "raw_slide": new_text,
            "verseTag": "1",
        }]
        return True

    return False


def _replace_custom_item_reference(service_data: list, marker_title: str, reference: str, scripture_text: str) -> bool:
    """
    Replace one custom service item (matched by header.title) with scripture content.
    Uses the existing first-line prompt from the holder slide when present.
    """
    marker = marker_title.strip().lower()

    for item in service_data:
        svc = item.get("serviceitem")
        if not isinstance(svc, dict):
            continue

        header = svc.get("header")
        if not isinstance(header, dict):
            continue

        if str(header.get("plugin", "")).lower() != "custom":
            continue

        title = str(header.get("title", "")).strip().lower()
        if title != marker:
            continue

        header["title"] = reference.strip()
        header["footer"] = [f"{reference.strip()} "]

        data = svc.get("data")
        if not isinstance(data, list) or not data:
            data = [{"title": "Reading", "raw_slide": "", "verseTag": "1"}]
            svc["data"] = data

        first = data[0] if isinstance(data[0], dict) else {}
        raw = str(first.get("raw_slide", "") or "")
        first_line = raw.splitlines()[0].strip() if raw else ""
        if first_line.lower().endswith("_holder"):
            first_line = ""

        if first_line and first_line.lower() != reference.strip().lower():
            new_raw = f"{first_line}\n{reference.strip()}\n\n{scripture_text.strip()}"
        else:
            new_raw = f"{reference.strip()}\n\n{scripture_text.strip()}"

        first["raw_slide"] = new_raw
        first.setdefault("title", (first_line or "Reading")[:30])
        first.setdefault("verseTag", "1")
        if not isinstance(data[0], dict):
            data[0] = first

        return True

    return False


def _fetch_scripture_verses(church: str, passage: str) -> list[tuple[int, int, str]]:
    """Fetch scripture verses as (chapter, verse, text) for a church Bible DB."""
    from scripts.utils.openlp import get_bible_db, connect_sqlite
    from scripts.utils.text_clean import full_scrub

    db_path = get_bible_db(church)
    conn = connect_sqlite(db_path)
    cur = conn.cursor()

    passage = _sanitize_scripture_reference_for_openlp(passage)
    m = re.match(r"^([1-3]?\s?[A-Za-z ]+)\s+(.+)$", passage)
    if not m:
        conn.close()
        raise ValueError(f"Unable to parse scripture reference: {passage}")

    book, remainder = m.group(1).strip(), m.group(2).strip()
    cur.execute("SELECT id FROM book WHERE name LIKE ?", (book,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise ValueError(f"Book '{book}' not found in DB")
    book_id = row["id"]

    verses_out: list[tuple[int, int, str]] = []
    last_chapter = None

    for segment in remainder.split(";"):
        segment = segment.strip()
        if not segment:
            continue

        for part in segment.split(","):
            part = part.strip()
            if not part:
                continue

            if ":" in part:
                chap_str, verse_expr = part.split(":", 1)
                chapter = int(chap_str.strip())
                last_chapter = chapter
            else:
                chapter = last_chapter
                verse_expr = part

            if chapter is None:
                conn.close()
                raise ValueError(f"Unable to determine chapter in scripture reference: {passage}")

            verse_expr = verse_expr.strip()
            if not verse_expr:
                continue

            if "-" in verse_expr:
                v1_str, v2_str = verse_expr.split("-", 1)
                v1, v2 = int(v1_str.strip()), int(v2_str.strip())
                cur.execute(
                    """
                    SELECT chapter, verse, text
                    FROM verse
                    WHERE book_id=? AND chapter=? AND verse BETWEEN ? AND ?
                    ORDER BY verse
                    """,
                    (book_id, chapter, v1, v2),
                )
            else:
                v = int(verse_expr)
                cur.execute(
                    """
                    SELECT chapter, verse, text
                    FROM verse
                    WHERE book_id=? AND chapter=? AND verse=?
                    """,
                    (book_id, chapter, v),
                )

            for r in cur.fetchall():
                verses_out.append((int(r["chapter"]), int(r["verse"]), full_scrub(r["text"] or "")))

    conn.close()
    return verses_out


def _sanitize_scripture_reference_for_openlp(reference: str) -> str:
        """
        Normalize scripture refs for OpenLP/DB parsing.

        Keeps the original reference for bulletin/Obsidian display, but strips
        verse-part letters immediately following numerals so int parsing works:
            - Genesis 12:1-4a -> Genesis 12:1-4
            - John 3:16b,17 -> John 3:16,17
        """
        cleaned = reference.replace("–", "-").replace("—", "-").strip()
        return re.sub(r"(?<=\d)[A-Za-z]+(?=(?:\s|$|[-,;]))", "", cleaned)


def _build_bible_slide_data(verses: list[tuple[int, int, str]]) -> list[dict]:
    """Build OpenLP-style slide rows from numbered scripture verses."""
    rows: list[dict] = []
    for chapter, verse, text in verses:
        prefix = f"{{su}}{chapter}:{verse}&nbsp;{{/su}}"
        raw = f"{prefix}{text}".strip()
        title = raw[:30]
        rows.append({
            "title": title,
            "raw_slide": raw,
            "verseTag": str(verse),
        })
    return rows


def _replace_bible_item_by_order(service_data: list, bible_order: int, reference: str, church: str) -> bool:
    """
    Replace nth bibles-plugin item in service_data with a new passage from the church DB.
    bible_order is 1-based.
    """
    bible_items = []
    for item in service_data:
        svc = item.get("serviceitem") if isinstance(item, dict) else None
        if not isinstance(svc, dict):
            continue
        header = svc.get("header")
        if not isinstance(header, dict):
            continue
        if str(header.get("plugin", "")).lower() == "bibles":
            bible_items.append(svc)

    if len(bible_items) < bible_order:
        return False

    svc = bible_items[bible_order - 1]
    header = svc.get("header", {})

    verses = _fetch_scripture_verses(church, reference)
    if not verses:
        raise ValueError(f"No verses returned for reference: {reference}")

    # Preserve Bible metadata from template holder.
    bmeta = {}
    try:
        bmeta = header.get("data", {}).get("bibles", [])[0]
    except Exception:
        bmeta = {}

    version = str(bmeta.get("version", "")).strip()
    copyright_text = str(bmeta.get("copyright", "")).strip()
    permissions = str(bmeta.get("permissions", "")).strip()
    meta_line = ", ".join([x for x in (version, copyright_text, permissions) if x])

    header["title"] = f"{reference} {meta_line}".strip()
    header["footer"] = [reference, meta_line] if meta_line else [reference]
    svc["data"] = _build_bible_slide_data(verses)
    return True


def _inject_custom_slides_into_openlp_service(osz_path: Path, church: str, master_md: str) -> None:
    """
    Inject matching existing custom slides into copied OpenLP service archives:
      - CtW for both churches
      - AoF for Elkton
    """
    from scripts.utils.placeholder import extract_block
    from scripts.utils.openlp import get_scripture_text

    ctw_ref = extract_block(master_md, "ctw_ref").strip()
    aof_ref = extract_block(master_md, "aof_ref").strip()
    scripture_ref_1 = extract_block(master_md, "scripture_ref_1").strip()
    scripture_ref_2 = extract_block(master_md, "scripture_ref_2").strip()
    scripture_ref_3 = extract_block(master_md, "scripture_ref_3").strip()

    service_data, payloads, infos = _load_service_data_from_osz(osz_path)
    changed = False

    if ctw_ref:
        ctw = _get_slide_text_for_prefix(church, f"CtW {ctw_ref}")
        if ctw:
            ctw_title, ctw_text = ctw
            if _replace_custom_item_text(service_data, "ctw_holder", ctw_title, ctw_text):
                changed = True
            else:
                logging.warning("CtW marker slide not found in %s service template", church)
        else:
            logging.warning("No CtW custom slide matched for %s: %s", church, ctw_ref)

    if church == "elkton" and aof_ref:
        aof = _get_slide_text_for_prefix(church, f"AoF p{aof_ref}")
        if aof:
            aof_title, aof_text = aof
            if _replace_custom_item_text(service_data, "aof_holder", aof_title, aof_text):
                changed = True
            else:
                logging.warning("AoF marker slide not found in elkton service template")
        else:
            logging.warning("No AoF custom slide matched for elkton: p%s", aof_ref)

    # Scripture updates:
    # 1) Prefer native OpenLP Bible plugin items by order (1, 2, 3)
    # 2) Fallback to custom holder slides if no bibles items are found
    bible_item_count = 0
    for item in service_data:
        svc = item.get("serviceitem") if isinstance(item, dict) else None
        if not isinstance(svc, dict):
            continue
        header = svc.get("header")
        if isinstance(header, dict) and str(header.get("plugin", "")).lower() == "bibles":
            bible_item_count += 1

    refs = ((1, scripture_ref_1), (2, scripture_ref_2), (3, scripture_ref_3))
    if bible_item_count >= 3:
        for idx, ref in refs:
            if not ref:
                continue
            try:
                if _replace_bible_item_by_order(service_data, idx, ref, church):
                    changed = True
                else:
                    logging.warning("Bible service item #%s not found in %s service template", idx, church)
            except Exception as e:
                logging.warning("Could not update Bible item #%s (%s) for %s: %s", idx, ref, church, e)
    else:
        for idx, ref in refs:
            if not ref:
                continue

            marker = f"scripture_ref_{idx}_holder"
            try:
                scripture_text = get_scripture_text(church, ref)
            except Exception as e:
                logging.warning("Could not load scripture %s for %s: %s", ref, church, e)
                continue

            if _replace_custom_item_reference(service_data, marker, ref, scripture_text):
                changed = True
            else:
                logging.warning("Scripture marker slide %s not found in %s service template", marker, church)

    if changed:
        _write_service_data_to_osz(osz_path, service_data, payloads, infos)


def copy_openlp_templates_for_each_church(
    *,
    master_path: Path | None = None,
    strict: bool = False,
    verbose: bool = False,
) -> Dict[str, Path]:
    """
    For each church, copy the OpenLP template file to a new file in the source directory
    with a date-based name, using config and file I/O utilities.
    Returns a dict mapping church name to the new file path.
    """
    from scripts.utils.config import config
    from scripts.utils.file_io import safe_mkdir
    from scripts.utils.file_io import read_text, write_text
    from scripts.utils.placeholder import extract_block

    init_logging(verbose=verbose)

    worship_dir = config.worship_dir

    if master_path is None:
        raise ValueError("master_path must be provided")
    if not master_path.exists():
        raise FileNotFoundError(f"Master.md not found at: {master_path}")

    master_md = read_text(master_path)
    communion = extract_block(master_md, "communion").strip() or "0"
    templates = PASTOR_TEMPLATE_BY_COMMUNION.get(communion)
    if not templates:
        raise ValueError(f"No templates for communion type: {communion}")

    date_slug = _get_date_slug(master_md)

    # Map church to OpenLP root and template key
    church_info = {
        "elkton": {
            "root": getattr(config, "elkton_root", None),
            "template_key": "elkton_openlp",
        },
        "lb": {
            "root": getattr(config, "lb_root", None),
            "template_key": "lb_openlp",
        },
    }

    output_paths = {}

    for church, info in church_info.items():
        openlp_root = info["root"]
        template_key = info["template_key"]
        template_name = templates.get(template_key)
        if not openlp_root or not openlp_root.exists():
            logging.warning(f"OpenLP root for {church} not found: {openlp_root}")
            continue
        if not template_name:
            logging.warning(f"No OpenLP template defined for {church} (communion {communion})")
            continue
        # OpenLP .osz files are in 'OpenLP Services' subdir
        service_dir = openlp_root / "OpenLP Services"
        src_path = service_dir / template_name
        if not src_path.exists():
            logging.warning(f"OpenLP template not found for {church}: {src_path}")
            continue
        # Output to the same subdirectory with 'Service--{date_slug}.osz' as the name
        out_name = f"Service--{date_slug}.osz"
        out_path = service_dir / out_name
        # Copy file (binary)
        with src_path.open("rb") as fsrc, out_path.open("wb") as fdst:
            fdst.write(fsrc.read())

        try:
            _inject_custom_slides_into_openlp_service(out_path, church, master_md)
        except Exception as e:
            logging.warning("Could not inject custom slides into %s: %s", out_path, e)

        output_paths[church] = out_path
        logging.info(f"Copied OpenLP template for {church} to {out_path}")

    return output_paths



# Combined function for both markdown and writer outputs
def build_markdown_and_writer_outputs(
    *,
    master_path: Path | None = None,
    templates_dir: Path | None = None,
    strict: bool = False,
    verbose: bool = False,
) -> Tuple[Dict[str, RenderResult], RenderResult, Dict[str, Path]]:
    """
    Generate:
      - Pastor/PastorC* speaker markdown files for both churches
      - Liturgist markdown file (unchanged)
      - Writer files for both churches

    Returns (speaker_results, liturgist_result, writer_results)
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
    templates = PASTOR_TEMPLATE_BY_COMMUNION.get(communion)
    if not templates:
        raise ValueError(f"No templates for communion type: {communion}")

    date_slug = _get_date_slug(master_md)

    speaker_results = {}
    writer_results = {}

    for church in ("elkton", "lb"):
        md_key = f"{church}_md"
        writer_key = f"{church}_writer"
        md_template_name = templates[md_key]
        writer_template_name = templates[writer_key]

        # Markdown
        md_template_path = templates_dir / md_template_name
        if not md_template_path.exists():
            raise FileNotFoundError(f"Speaker template not found: {md_template_path}")
        md_template = read_text(md_template_path)
        md_rendered, md_warnings = render_markdown_template(
            md_template,
            master_md,
            strict=strict,
        )
        md_out = worship_dir / f"{Path(md_template_name).stem}--{date_slug}.md"
        write_text(md_out, md_rendered)
        speaker_results[church] = RenderResult(md_out, md_warnings)

        # Writer
        # Use build_librewriter_outputs to get year_dir
        writer_info = build_librewriter_outputs(master_path=master_path, strict=strict, verbose=verbose)
        year_dir = writer_info["year_dir"]
        writer_template_path = year_dir / writer_template_name
        if not writer_template_path.exists():
            logging.warning(f"Writer template not found: {writer_template_path}")
            writer_results[church] = None
        else:
            # Replace placeholders in content.xml of the Writer template
            import zipfile
            from io import BytesIO

            writer_out = year_dir / f"{Path(writer_template_name).stem}--{date_slug}.odt"
            with zipfile.ZipFile(writer_template_path, 'r') as zin:
                with zipfile.ZipFile(writer_out, 'w') as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)
                        if item.filename == 'content.xml':
                            # Replace placeholders in content.xml
                            text = data.decode('utf-8')
                            # Use the same placeholder replacement as markdown
                            # Build lookup for placeholders used in content.xml
                            placeholders = extract_placeholders_used(text)
                            lookup = build_master_lookup(master_md, placeholders)
                            for name in placeholders:
                                value = lookup.get(name, "")
                                value = _format_writer_placeholder_value(name, value)
                                text = text.replace(f"{{{name}}}", value)
                            data = text.encode('utf-8')
                        zout.writestr(item, data)
            writer_results[church] = writer_out

    # Liturgist (unchanged)
    liturgist_template_path = templates_dir / DEFAULT_LITURGIST_TEMPLATE
    if not liturgist_template_path.exists():
        raise FileNotFoundError(f"Liturgist template not found: {liturgist_template_path}")
    liturgist_template = read_text(liturgist_template_path)
    liturgist_rendered, liturgist_warnings = render_markdown_template(
        liturgist_template,
        master_md,
        strict=strict,
    )
    liturgist_out = _liturgist_output_path(worship_dir, date_slug)
    write_text(liturgist_out, liturgist_rendered)
    liturgist_result = RenderResult(liturgist_out, liturgist_warnings)

    # Log warnings
    for church, result in speaker_results.items():
        for w in result.warnings:
            logging.warning(f"[Speaker {church}] %s", w)
    for w in liturgist_warnings:
        logging.warning("[Liturgist] %s", w)

    return speaker_results, liturgist_result, writer_results
# scripts/publish.py



import logging
import re
import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Tuple
from datetime import datetime

from scripts.utils.config import config
from scripts.utils.file_io import read_text, write_text, safe_mkdir
from scripts.utils.logging_utils import init_logging
from scripts.utils.placeholder import extract_block
from scripts.utils.text_clean import clean_text, clean_markdown


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
        raw_value = extract_block(master_md, name)
        value = _strip_noncontent_lines_for_templates(raw_value)
        value = _format_song_id_for_templates(name, value)
        value = _format_multiline_fields_for_templates(name, value)
        lookup[name] = value
    return lookup


def _strip_noncontent_lines_for_templates(value: str) -> str:
    """
    Remove markdown structural/comment-only lines from extracted blocks.

    Examples removed:
      - '# Music'
      - '%% b = blue, r = red, k = black %%'
      - '<!-- comment -->'
    """
    if not value:
        return value

    lines = value.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    kept: list[str] = []
    for line in lines:
        stripped = line.strip()

        if re.match(r"^#{1,6}\s+\S", stripped):
            continue
        if re.match(r"^%%.*%%$", stripped):
            continue
        if re.match(r"^<!--.*-->$", stripped):
            continue

        kept.append(line)

    return "\n".join(kept).strip("\n")


def _format_song_id_for_templates(placeholder_name: str, value: str) -> str:
    """
    Format song IDs for bulletin/Obsidian template rendering.

        Examples:
            - b42   -> blue #42
            - r378  -> red #378
            - k91   -> black #91

    Leaves non-song-id placeholders unchanged.
    """
    if not (placeholder_name.startswith("song_") and "_id_" in placeholder_name):
        return value

    base = value.split("%%", 1)[0].strip()
    m = re.match(r"^([brkBRK])(\d+)(.*)$", base)
    if not m:
        return value

    prefix = m.group(1).lower()
    color = {"b": "blue", "r": "red", "k": "black"}[prefix]
    number = m.group(2)
    remainder = (m.group(3) or "").strip()
    formatted = f"{color} #{number}"
    return f"{formatted} {remainder}".strip()


def _format_multiline_fields_for_templates(placeholder_name: str, value: str) -> str:
    """
    For selected placeholders, keep one newline between paragraphs while
    collapsing wrapped lines inside each paragraph to spaces.

    This preserves markdown paragraph intent for publishing outputs.
    """
    if not _requires_single_newline_paragraphs(placeholder_name):
        return value

    normalized = value.replace("\r\n", "\n").replace("\r", "\n")
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n+", normalized) if p.strip()]
    collapsed = [re.sub(r"\s+", " ", p) for p in paragraphs]
    return "\n".join(collapsed)


def _requires_single_newline_paragraphs(placeholder_name: str) -> bool:
    return (
        placeholder_name.startswith("ctw_xml_")
        or placeholder_name.startswith("announce_events_")
        or placeholder_name.startswith("announce_event_")
        or placeholder_name.startswith("announce_detail_")
        or placeholder_name.startswith("announce_details_")
        or placeholder_name in {"benediction_text", "benediction_text_alt"}
    )


def _format_writer_placeholder_value(placeholder_name: str, value: str) -> str:
    """
    Writer/ODT content.xml requires explicit line-break tags in text runs.
    Apply only to placeholders where we intentionally retain paragraph breaks.
    """
    if not _requires_single_newline_paragraphs(placeholder_name):
        return value
    return value.replace("\n", "<text:line-break/>")


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
    templates = PASTOR_TEMPLATE_BY_COMMUNION.get(communion)
    template_name = (templates or {}).get("elkton_md", "Pastor.md")

    speaker_template_path = templates_dir / template_name
    liturgist_template_path = templates_dir / DEFAULT_LITURGIST_TEMPLATE

    if not speaker_template_path.exists():
        raise FileNotFoundError(f"Speaker template not found: {speaker_template_path}")
    if not liturgist_template_path.exists():
        raise FileNotFoundError(f"Liturgist template not found: {liturgist_template_path}")

    date_slug = _get_date_slug(master_md)

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
# LibreOffice Writer output builder (stub)
# -----------------------------------------------------------------------------

def build_librewriter_outputs(*, master_path: Path | None = None, strict: bool = False, verbose: bool = False):
    """
    Access the appropriate subdirectory under bulletin_dir based on the year in {cal_date}.
    This will be expanded to generate Writer files from templates in that subdirectory.
    """
    from scripts.utils.config import config
    from scripts.utils.file_io import read_text

    if master_path is None:
        raise ValueError("master_path must be provided")
    if not master_path.exists():
        raise FileNotFoundError(f"Master.md not found at: {master_path}")

    master_md = read_text(master_path)
    # Extract the calendar date from the master file
    from scripts.utils.placeholder import extract_block
    cal_date = extract_block(master_md, "cal_date").strip()
    if not cal_date:
        raise ValueError("{cal_date} not found in master file")


    # Try to parse the year from cal_date (supporting multiple formats)
    year = None
    date_formats = [
        "%Y-%m-%d",
        "%B %d, %Y",  # e.g., January 25, 2026
        "%b %d, %Y",  # e.g., Jan 25, 2026
        "%m/%d/%Y",
        "%d-%b-%Y",
        "%d %B %Y",
    ]
    for fmt in date_formats:
        try:
            year = datetime.strptime(cal_date, fmt).year
            break
        except Exception:
            continue
    if not year:
        # fallback: try to extract a 4-digit year
        import re
        m = re.search(r"(20\\d{2})", cal_date)
        if m:
            year = m.group(1)
        else:
            raise ValueError(f"Could not determine year from cal_date: {cal_date}")

    bulletin_dir = getattr(config, "bulletin_dir", None)
    if bulletin_dir is None or not bulletin_dir:
        raise ValueError("bulletin_dir is not configured in weekly_config.ini")

    year_dir = Path(bulletin_dir) / str(year)
    if not year_dir.exists():
        raise FileNotFoundError(f"Bulletin year directory does not exist: {year_dir}")

    # List all files in the year directory (LibreOffice templates and Writer docs)
    from scripts.utils.file_io import list_files
    files = list(year_dir.glob("*.ott")) + list(year_dir.glob("*.odt"))

    # For now, just return the directory and file list (stub for future expansion)
    return {"year_dir": year_dir, "files": files}


# -----------------------------------------------------------------------------
# CLI entry point
# -----------------------------------------------------------------------------


if __name__ == "__main__":
    args = parse_args()

    worship_dir = config.worship_dir
    master_path = worship_dir / args.master

    speaker_results, liturgist_result, writer_results = build_markdown_and_writer_outputs(
        master_path=master_path,
        strict=args.strict,
        verbose=args.verbose,
    )

    for church, result in speaker_results.items():
        print(f"Speaker file for {church}: {result.output_path.name}")
        if result.warnings:
            print(f"  Warnings for {church}:")
            for w in result.warnings:
                print("   -", w)

    print(f"Liturgist file: {liturgist_result.output_path.name}")
    if liturgist_result.warnings:
        print("\nLiturgist warnings:")
        for w in liturgist_result.warnings:
            print(" -", w)

    for church, writer_path in writer_results.items():
        if writer_path:
            print(f"Writer file for {church}: {writer_path.name}")
        else:
            print(f"Writer template for {church} not found.")

    # Call OpenLP template copier and print results
    openlp_results = copy_openlp_templates_for_each_church(
        master_path=master_path,
        strict=args.strict,
        verbose=args.verbose,
    )
    for church, out_path in openlp_results.items():
        print(f"OpenLP file for {church}: {out_path.name}")
