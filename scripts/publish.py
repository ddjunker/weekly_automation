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

from scripts.utils.config import config
from scripts.utils.file_io import read_text, write_text
from scripts.utils.logging_utils import init_logging
from scripts.utils.placeholder import extract_block, has_placeholder
from scripts.utils.text_clean import clean_text, clean_markdown

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


def _get_slide_text_for_prefix(church: str, prefix: str, exact: bool = False) -> tuple[str, str] | None:
    """
    Find a custom slide whose title starts with prefix and return (title, cleaned_text).
    When exact=True, only a slide whose normalized title equals prefix is accepted.
    """
    from scripts.utils.openlp import list_custom_slides, load_custom_slide
    from scripts.utils.text_clean import clean_text

    prefix_clean = clean_text(prefix).lower().replace(",", "")
    slides = list_custom_slides(church)

    if exact:
        matches = [s for s in slides if clean_text(s["title"]).lower().replace(",", "") == prefix_clean]
        chosen = matches[0] if matches else None
    else:
        matches = [s for s in slides if clean_text(s["title"]).lower().replace(",", "").startswith(prefix_clean)]
        if not matches:
            return None
        # Prefer exact normalized title match if available.
        chosen = next(
            (s for s in matches if clean_text(s["title"]).lower().replace(",", "") == prefix_clean),
            matches[0],
        )

    if not chosen:
        return None

    slide = load_custom_slide(church, chosen["uuid"])
    if not slide or not isinstance(slide.get("raw_text"), str) or not slide["raw_text"].strip():
        return None

    from scripts.utils.openlp import extract_openlp_slide_text
    text = extract_openlp_slide_text(slide["raw_text"])
    if not text:
        return None

    return slide["raw_title"], text


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
        header["data"] = {"title": new_title, "credits": ""}
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


def _replace_custom_item_second_line(service_data: list, marker_title: str, new_second_line: str) -> bool:
    """
    For a custom slide matched by header.title, keep line 1 of raw_slide unchanged
    and replace line 2 with new_second_line.
    """
    marker = marker_title.strip().lower()
    replacement = (new_second_line or "").strip()

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

        data = svc.get("data")
        if not isinstance(data, list) or not data:
            return False
        first = data[0] if isinstance(data[0], dict) else {}

        raw = str(first.get("raw_slide", "") or "")
        lines = raw.splitlines()
        if not lines:
            return False

        # Keep line 1, replace line 2, preserve any remaining lines.
        if len(lines) == 1:
            lines.append(replacement)
        else:
            lines[1] = replacement

        first["raw_slide"] = "\n".join(lines)
        if not isinstance(data[0], dict):
            data[0] = first

        return True

    return False


def _replace_offertory_prayer_content_and_credits(service_data: list, body_text: str, credit_text: str) -> bool:
    """
    Update the custom slide titled 'Offertory Prayer':
            - Rebuild slide data from scratch to avoid retaining stale template text
            - Replace body text with body_text
      - Update header.footer so credits project from offertory_credit
    """
    body = (body_text or "").strip()
    if not body:
        return False

    credits_raw = (credit_text or "").replace("\r\n", "\n").replace("\r", "\n")
    credit_lines = [line.strip() for line in credits_raw.split("\n") if line.strip()]

    for item in service_data:
        svc = item.get("serviceitem")
        if not isinstance(svc, dict):
            continue

        header = svc.get("header")
        if not isinstance(header, dict):
            continue

        if str(header.get("plugin", "")).strip().lower() != "custom":
            continue

        title = str(header.get("title", "")).strip().lower()
        if title != "offertory prayer":
            continue

        svc["data"] = [{
            "title": "Offertory Prayer",
            "raw_slide": f"{{h}}Offertory Prayer{{/h}}\n\n{body}\n",
            "verseTag": "1",
        }]

        footer_title = "Offertory Prayer"
        header["title"] = footer_title
        header["footer"] = [footer_title, *credit_lines] if credit_lines else [footer_title]
        return True

    return False


def _format_benediction_slides_for_openlp(benediction_text: str) -> list[str] | None:
    """
    Return 3 deterministic Benediction slide bodies:
      1) Paragraph 1 flattened, bounded by yellow tags
      2) Paragraph 2 flattened, bounded by yellow tags
      3) Fixed congregational response
    """
    normalized = (benediction_text or "").replace("\r\n", "\n").replace("\r", "\n")
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n+", normalized) if p.strip()]
    flattened = [re.sub(r"\s+", " ", p) for p in paragraphs]

    if len(flattened) < 2:
        return None

    slide_1 = f"{{y}}{flattened[0]}{{/y}}"
    slide_2 = f"{{y}}{flattened[1]}{{/y}}"
    slide_3 = (
        "{p}{y}We go in peace to love and serve the Lord,{/y}{/p}\n"
        "{it}All:{/it} {st}In the name of Christ. Amen.{/st}"
    )
    return [slide_1, slide_2, slide_3]


def _replace_benediction_custom_slide(service_data: list, benediction_text: str) -> bool:
    """
    Update custom slide titled 'Benediction':
      - Keep the first raw_slide line unchanged (heading/markup)
      - Replace body with formatted benediction content
    """
    slides = _format_benediction_slides_for_openlp(benediction_text)
    if not slides:
        return False

    for item in service_data:
        svc = item.get("serviceitem")
        if not isinstance(svc, dict):
            continue

        header = svc.get("header")
        if not isinstance(header, dict):
            continue

        if str(header.get("plugin", "")).strip().lower() != "custom":
            continue

        title = str(header.get("title", "")).strip().lower()
        if title != "benediction":
            continue

        svc["data"] = [{
            "title": "Benediction",
            "raw_slide": slides[0],
            "verseTag": "1",
        }, {
            "title": "Benediction",
            "raw_slide": slides[1],
            "verseTag": "2",
        }, {
            "title": "Benediction",
            "raw_slide": slides[2],
            "verseTag": "3",
        }]

        header["title"] = "Benediction"

        return True

    return False


def _remove_service_items_by_marker(service_data: list, plugin: str, marker_title: str) -> int:
    """Remove all service items matching plugin + header.title (case-insensitive). Returns count removed."""
    plugin = (plugin or "").strip().lower()
    marker = (marker_title or "").strip().lower()
    if not plugin or not marker:
        return 0

    kept = []
    removed = 0
    for item in service_data:
        svc = item.get("serviceitem") if isinstance(item, dict) else None
        header = svc.get("header") if isinstance(svc, dict) else None
        if isinstance(header, dict):
            item_plugin = str(header.get("plugin", "")).strip().lower()
            item_title = str(header.get("title", "")).strip().lower()
            if item_plugin == plugin and item_title == marker:
                removed += 1
                continue
        kept.append(item)

    service_data[:] = kept
    return removed


def _replace_song_item_by_marker(
    service_data: list,
    marker_title: str,
    new_title: str,
    lyrics_text: str,
    *,
    slides: list[dict] | None = None,
    authors: str = "",
    copyright_text: str = "",
    ccli_number: str = "",
    hymnal: str = "",
    entry: str = "",
    search_title: str = "",
    alternate_title: str = "",
) -> bool:
    """
    Replace one Songs-plugin service item (matched by header.title) with new song text
    and refresh song metadata in header.data + header.footer.
    """
    marker = (marker_title or "").strip().lower()
    if not marker:
        return False

    if slides is None:
        # Fall back: split plain-text lyrics on blank lines, label V1 V2 ...
        normalized = (lyrics_text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
        parts = [p.strip() for p in re.split(r"\n\s*\n+", normalized) if p.strip()]
        if not parts:
            parts = [normalized] if normalized else []
        slides = [
            {
                "title": (chunk[:30] or new_title[:30] or "Song")[:30],
                "raw_slide": chunk,
                "verseTag": f"V{idx}",
            }
            for idx, chunk in enumerate(parts, start=1)
        ]

    for item in service_data:
        svc = item.get("serviceitem")
        if not isinstance(svc, dict):
            continue

        header = svc.get("header")
        if not isinstance(header, dict):
            continue

        if str(header.get("plugin", "")).strip().lower() != "songs":
            continue

        title = str(header.get("title", "")).strip().lower()
        if title != marker:
            continue

        header["title"] = new_title
        # Preserve configured CCLI license line from template footer when present.
        existing_footer = header.get("footer")
        existing_footer_lines = existing_footer if isinstance(existing_footer, list) else []
        license_line = next(
            (
                str(line).strip()
                for line in existing_footer_lines
                if isinstance(line, str) and str(line).strip().lower().startswith("ccli license:")
            ),
            "",
        )

        footer_lines = [new_title]

        # authors from DB are comma-separated; footer display uses "and"
        authors_clean = (authors or "").strip()
        if authors_clean:
            authors_display = " and ".join(a.strip() for a in authors_clean.split(",") if a.strip())
            footer_lines.append(f"Written by: {authors_display}")

        copyright_clean = (copyright_text or "").strip()
        if copyright_clean:
            # Avoid double-prefix when DB value already starts with © or (c)
            if copyright_clean.startswith("©") or copyright_clean.lower().startswith("(c)"):
                footer_lines.append(copyright_clean)
            else:
                footer_lines.append(f"© {copyright_clean}")

        hymnal_clean = (hymnal or "").strip()
        entry_clean = (entry or "").strip()
        if hymnal_clean and entry_clean:
            footer_lines.append(f"{hymnal_clean} #{entry_clean}")
        elif hymnal_clean:
            footer_lines.append(hymnal_clean)

        ccli_song = (ccli_number or "").strip()
        if ccli_song:
            footer_lines.append(f"CCLI Song # {ccli_song}")

        if license_line:
            footer_lines.append(license_line)

        header["footer"] = footer_lines

        # Use search_title (e.g. "as the deer@") for hdata["title"] so OpenLP
        # can match the song in its database; fall back to display title.
        data_title = (search_title or "").strip() or new_title
        hdata = header.get("data")
        if isinstance(hdata, dict):
            hdata["title"] = data_title
        else:
            header["data"] = {"title": data_title}
            hdata = header["data"]

        hdata["alternate_title"] = (alternate_title or "").strip()
        hdata["authors"] = authors_clean
        hdata["copyright"] = copyright_clean
        hdata["ccli_number"] = (ccli_number or "").strip()

        # Update audit field: [title, [authors_list], copyright, ccli_number]
        # OpenLP uses this to display/refresh the footer in the service manager.
        authors_list = [a.strip() for a in authors_clean.split(",") if a.strip()] if authors_clean else []
        header["audit"] = [
            new_title,
            authors_list,
            copyright_clean,
            (ccli_number or "").strip(),
        ]

        svc["data"] = slides
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
    from scripts.utils.openlp import get_scripture_text

    ctw_ref = extract_block(master_md, "ctw_ref").strip()
    ctw_ref_church = extract_block(master_md, f"ctw_ref_{church}").strip()
    aof_ref = extract_block(master_md, "aof_ref").strip()
    scripture_ref_1 = extract_block(master_md, "scripture_ref_1").strip()
    scripture_ref_2 = extract_block(master_md, "scripture_ref_2").strip()
    scripture_ref_3 = extract_block(master_md, "scripture_ref_3").strip()
    offertory_source_text = extract_block(master_md, "offertory_source_text").strip()
    offertory_credit = extract_block(master_md, "offertory_credit").strip()
    benediction_text = extract_block(master_md, "benediction_text").strip()

    # Songs (OpenLP markers live in the service templates)
    song_opening_title = extract_block(master_md, f"song_opening_title_{church}").strip()
    song_opening_id = extract_block(master_md, f"song_opening_id_{church}").strip()
    song_middle_title = extract_block(master_md, f"song_middle_title_{church}").strip()
    song_middle_id = extract_block(master_md, f"song_middle_id_{church}").strip()
    song_closing_title = extract_block(master_md, f"song_closing_title_{church}").strip()
    song_closing_id = extract_block(master_md, f"song_closing_id_{church}").strip()

    service_data, payloads, infos = _load_service_data_from_osz(osz_path)
    changed = False

    if ctw_ref_church:
        ctw = _get_slide_text_for_prefix(church, ctw_ref_church, exact=True)
        if ctw:
            ctw_title, ctw_text = ctw
            if _replace_custom_item_text(service_data, "ctw_holder", ctw_title, ctw_text):
                changed = True
            else:
                logging.warning("CtW marker slide not found in %s service template", church)
        else:
            logging.warning("No CtW custom slide matched for %s: %s", church, ctw_ref_church)

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

    # Song updates (marker-based)
    try:
        if _inject_songs_into_openlp_service_data(
            service_data,
            church,
            opening_title=song_opening_title,
            opening_id=song_opening_id,
            middle_title=song_middle_title,
            middle_id=song_middle_id,
            closing_title=song_closing_title,
            closing_id=song_closing_id,
        ):
            changed = True
    except Exception as e:
        logging.warning("Could not inject songs into %s (%s): %s", osz_path, church, e)

    # Reading custom slides: keep first line, replace second line with scripture reference
    reading_map = (
        ("Reading: 1", scripture_ref_1),
        ("Reading: 2", scripture_ref_2),
        ("Reading: Gospel", scripture_ref_3),
    )
    for reading_title, reading_ref in reading_map:
        if not reading_ref:
            continue
        if _replace_custom_item_second_line(service_data, reading_title, reading_ref.strip()):
            changed = True

    if _replace_offertory_prayer_content_and_credits(
        service_data,
        offertory_source_text,
        offertory_credit,
    ):
        changed = True

    if _replace_benediction_custom_slide(service_data, benediction_text):
        changed = True

    if changed:
        _write_service_data_to_osz(osz_path, service_data, payloads, infos)


def _inject_songs_into_openlp_service_data(service_data: list, church: str, *,
                                          opening_title: str, opening_id: str,
                                          middle_title: str, middle_id: str,
                                          closing_title: str, closing_id: str) -> bool:
    """Return True if any song items were updated/removed using ID-driven slot rules."""
    from scripts.utils.text_clean import xml_to_text
    from scripts.utils.openlp import build_song_index, load_song
    from scripts.music_gather import hymnal_from_prefix

    changed = False
    church_key = (church or "").strip().lower()

    def warn_song_missing(slot_marker: str, message: str) -> None:
        msg = f"[OpenLP song missing] {church_key}:{slot_marker} - {message}"
        print(msg)

    def normalize_song_id(raw_id: str) -> str:
        if raw_id is None:
            return ""
        base = str(raw_id).split("%%", 1)[0]
        return base.strip()

    def parse_hymnal_id(raw_id: str) -> tuple[str, str] | None:
        """Valid hymnal IDs: b/r/k immediately followed by 1-4 numerals."""
        m = re.match(r"^([brkBRK])(\d{1,4})$", raw_id)
        if not m:
            return None
        return m.group(1).lower(), m.group(2)

    def intro_prefix_for(title: str, prefix: str | None, entry: str | None) -> str | None:
        """
        Determine custom-slide prefix for intro holder replacement.
        - video: P901 <title>
        - k:     F####
        - r:     U###
        - b:     H### (lb) / S### (elkton)
        """
        if prefix is None or entry is None:
            clean_title = (title or "").strip()
            return f"P901 {clean_title}" if clean_title else None

        if prefix == "k":
            return f"F{entry.zfill(4)}"
        if prefix == "r":
            return f"U{entry.zfill(3)}"
        if prefix == "b":
            head = "H" if church_key == "lb" else "S"
            return f"{head}{entry.zfill(3)}"
        return None

    def apply_intro_from_custom_slide(holder_marker: str, title: str, prefix: str | None, entry: str | None) -> bool:
        intro_prefix = intro_prefix_for(title, prefix, entry)
        if not intro_prefix:
            return False
        slide = _get_slide_text_for_prefix(church_key, intro_prefix)
        if not slide:
            logging.warning("No intro custom slide matched for %s using prefix %r", holder_marker, intro_prefix)
            return False
        slide_title, slide_text = slide
        return _replace_custom_item_text(service_data, holder_marker, slide_title, slide_text)

    song_index = None

    def try_replace_song(song_marker: str, title: str, prefix: str, entry: str) -> bool:
        nonlocal song_index
        hymnal = hymnal_from_prefix(church_key, prefix)
        if not hymnal:
            warn_song_missing(song_marker, f"unknown hymnal prefix {prefix!r}")
            return False

        if song_index is None:
            song_index = build_song_index(church_key)

        uuid = song_index.get((hymnal, str(entry)))
        if not uuid:
            warn_song_missing(song_marker, f"not found in {hymnal} for id {prefix}{entry}")
            return False

        song = load_song(church_key, int(uuid))
        if not song:
            warn_song_missing(song_marker, f"UUID {uuid} not found")
            return False

        lyrics_xml = str(song.get("lyrics") or "").strip()
        if not lyrics_xml:
            warn_song_missing(song_marker, f"UUID {uuid} has no lyrics")
            return False

        from scripts.utils.openlp import parse_song_xml
        verse_order = str(song.get("verse_order") or "").strip()
        pre_slides = parse_song_xml(lyrics_xml, verse_order)
        lyrics_text = "" if pre_slides is not None else xml_to_text(lyrics_xml)
        new_title = str(song.get("raw_title") or song.get("title") or title or f"{hymnal} {entry}").strip()
        return _replace_song_item_by_marker(
            service_data,
            song_marker,
            new_title,
            lyrics_text,
            slides=pre_slides,
            authors=str(song.get("raw_authors") or song.get("authors") or "").strip(),
            copyright_text=str(song.get("raw_copyright") or song.get("copyright") or "").strip(),
            ccli_number=str(song.get("raw_ccli_number") or song.get("ccli_number") or "").strip(),
            hymnal=str(hymnal or "").strip(),
            entry=str(entry or "").strip(),
            search_title=str(song.get("search_title") or "").strip(),
            alternate_title=str(song.get("alternate_title") or "").strip(),
        )

    slots = (
        ("opening", opening_title, opening_id),
        ("middle", middle_title, middle_id),
        ("closing", closing_title, closing_id),
    )

    for slot_name, slot_title, slot_id_raw in slots:
        holder_marker = f"song_{slot_name}_holder"
        song_marker = f"song_{slot_name}"
        slot_id = normalize_song_id(slot_id_raw)

        # tbd means remove both intro + song placeholders for this slot
        if slot_id.lower() == "tbd":
            removed_custom = _remove_service_items_by_marker(service_data, "custom", holder_marker)
            removed_song = _remove_service_items_by_marker(service_data, "songs", song_marker)
            if removed_custom or removed_song:
                changed = True
            continue

        parsed = parse_hymnal_id(slot_id)
        if parsed is None:
            # Video/other non-hymnal ID: pick intro custom slide by P901 <title>, leave song placeholder unchanged.
            if apply_intro_from_custom_slide(holder_marker, slot_title, None, None):
                changed = True
            continue

        prefix, entry = parsed

        # For hymnal IDs, intro slide selection is based on prefix+number convention.
        if apply_intro_from_custom_slide(holder_marker, slot_title, prefix, entry):
            changed = True

        # Replace actual song item when resolvable; if not found, leave placeholder item unchanged.
        if try_replace_song(song_marker, slot_title, prefix, entry):
            changed = True

    return changed


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
    run_markdown: bool = True,
    run_writer: bool = True,
) -> Tuple[Dict[str, RenderResult], RenderResult | None, Dict[str, Path]]:
    """
    Generate any combination of:
      - Pastor/PastorC* speaker markdown files for both churches
      - Liturgist markdown file
      - Writer (ODT) files for both churches

    run_markdown / run_writer can be set False to skip those outputs.
    Returns (speaker_results, liturgist_result, writer_results);
    omitted segments return empty dicts / None.
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

    year_dir = None
    if run_writer:
        writer_info = build_librewriter_outputs(master_path=master_path, strict=strict, verbose=verbose)
        year_dir = writer_info["year_dir"]

    speaker_results = {}
    writer_results = {}

    for church in ("elkton", "lb"):
        md_template_name = templates[f"{church}_md"]

        if run_markdown:
            md_template_path = templates_dir / md_template_name
            if not md_template_path.exists():
                raise FileNotFoundError(f"Speaker template not found: {md_template_path}")
            md_template = read_text(md_template_path)
            lb_omit_overrides = {"reader": "nobody"} if church == "lb" else {}
            md_rendered, md_warnings = render_markdown_template(
                md_template,
                master_md,
                strict=strict,
                omit_overrides=lb_omit_overrides,
            )
            md_out = worship_dir / f"{Path(md_template_name).stem}--{date_slug}.md"
            write_text(md_out, md_rendered)
            speaker_results[church] = RenderResult(md_out, md_warnings)

        if run_writer:
            assert year_dir is not None
            writer_template_name = templates[f"{church}_writer"]
            writer_template_path = year_dir / writer_template_name
            if not writer_template_path.exists():
                logging.warning(f"Writer template not found: {writer_template_path}")
                writer_results[church] = None
            else:
                writer_out = year_dir / f"{Path(writer_template_name).stem}--{date_slug}.odt"
                with zipfile.ZipFile(writer_template_path, 'r') as zin:
                    with zipfile.ZipFile(writer_out, 'w') as zout:
                        for item in zin.infolist():
                            data = zin.read(item.filename)
                            if item.filename == 'content.xml':
                                text = data.decode('utf-8')
                                placeholders = extract_placeholders_used(text)
                                lookup = build_master_lookup(master_md, placeholders)
                                for name in placeholders:
                                    if name not in lookup:
                                        continue
                                    value = lookup[name]
                                    # In XML, "omit" and "delete" both collapse to empty string
                                    # (line-deletion would destroy ODT XML structure)
                                    if value.strip().lower() in ("omit", "delete"):
                                        value = ""
                                    if name in _PARA_SPLIT_PLACEHOLDERS:
                                        value = value.replace("&", "&amp;")
                                        text = _expand_writer_paragraph(text, name, value)
                                    else:
                                        value = _format_writer_placeholder_value(name, value)
                                        text = text.replace(f"{{{name}}}", value)
                                data = text.encode('utf-8')
                            zout.writestr(item, data)
                writer_results[church] = writer_out

    liturgist_result = None
    if run_markdown:
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
        for w in liturgist_warnings:
            logging.warning("[Liturgist] %s", w)

    for church, result in speaker_results.items():
        for w in result.warnings:
            logging.warning(f"[Speaker {church}] %s", w)

    return speaker_results, liturgist_result, writer_results


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
    segment = parser.add_mutually_exclusive_group()
    segment.add_argument(
        "-o", "--openlp",
        action="store_true",
        help="Update OpenLP files only",
    )
    segment.add_argument(
        "-w", "--writer",
        action="store_true",
        help="Update Writer (ODT) files only",
    )
    segment.add_argument(
        "-m", "--markdown",
        action="store_true",
        help="Update markdown speaker/liturgist files only",
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
        if not has_placeholder(master_md, name):
            continue  # not in master → leave template placeholder unchanged
        raw_value = extract_block(master_md, name)
        value = _strip_noncontent_lines_for_templates(raw_value)
        value = _format_aof_ref_for_templates(name, value)
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


def _format_aof_ref_for_templates(placeholder_name: str, value: str) -> str:
    """
    Normalize AoF references for bulletin/Obsidian template rendering.

    Example:
        - 38top -> 38 top
    """
    if placeholder_name != "aof_ref":
        return value

    base = value.split("%%", 1)[0].strip()
    m = re.match(r"^(\d+)([A-Za-z].*)$", base)
    if not m:
        return value

    return f"{m.group(1)} {m.group(2).strip()}".strip()


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


# Characters that should stay attached to preceding punctuation (no line break inserted between)
_CLOSING_QUOTES = "\u201d\u2019\"')]}"


def _requires_reading_reflow(placeholder_name: str) -> bool:
    """Return True for placeholders whose values should get punctuation-driven
    line breaks for easier public/pastoral reading."""
    return (
        re.match(r"^ctw_xml_(elkton|lb)$", placeholder_name) is not None
        or re.match(r"^scripture_ref_\d+_(elkton|lb)$", placeholder_name) is not None
        or placeholder_name in {"offertory_source_text", "benediction_text", "benediction_text_alt"}
    )


# Liturgical cue labels whose colon should never trigger a line break.
# Stored as (literal_string, placeholder_token) pairs.
_LITURGICAL_CUES = [
    ("Leader: ", "\x00LEAD\x00"),
    ("People: ", "\x00PEOP\x00"),
    ("L: ",      "\x00CL\x00"),
    ("P: ",      "\x00CP\x00"),
]


def _reflow_for_public_reading(placeholder_name: str, value: str) -> str:
    """
    Apply punctuation-driven line breaks for pastoral/public reading scripts.

    Rules (ported from md_clean.py):
      - After .  ?  !  (optionally followed by a closing/grouping char) → blank line
      - After :  ;  ,  (optionally followed by a closing/grouping char) → single line break
      - Closing quotes and grouping chars )]}" stay attached to the preceding punctuation.
      - Liturgical cue labels (Leader:, People:, L:, P:) are exempt from colon breaks.
      - 3+ consecutive newlines collapsed to 2.

    Applied per-paragraph (blank-line delimited) so multi-section blocks
    (e.g. L:/P: exchanges) keep their structural separation.
    """
    if not _requires_reading_reflow(placeholder_name) or not value.strip():
        return value

    normalized = value.replace("\r\n", "\n").replace("\r", "\n")
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n+", normalized) if p.strip()]

    cq = re.escape(_CLOSING_QUOTES)
    parts = []
    for para in paragraphs:
        t = re.sub(r"\s+", " ", para.strip())
        # Remove stray spaces before punctuation
        t = re.sub(r"\s+([,;:!?.])", lambda m: m.group(1), t)
        # Protect liturgical cue labels so their colons are not broken
        for orig, token in _LITURGICAL_CUES:
            t = t.replace(orig, token)
        # After . ? ! → blank line
        t = re.sub(
            rf"([.!?])\s*([{cq}]?)\s*",
            lambda m: f"{m.group(1)}{m.group(2)}\n\n",
            t,
        )
        # After : ; , → single line break
        t = re.sub(
            rf"([,;:])\s*([{cq}]?)\s*",
            lambda m: f"{m.group(1)}{m.group(2)}\n",
            t,
        )
        # Restore liturgical cue labels
        for orig, token in _LITURGICAL_CUES:
            t = t.replace(token, orig)
        t = re.sub(r"\n{3,}", "\n\n", t).strip()
        parts.append(t)

    return "\n\n".join(parts)


# Placeholders whose lines should each become a separate <text:p> element rather
# than a single paragraph with <text:line-break/> tags between them.
_PARA_SPLIT_PLACEHOLDERS = frozenset({
    "ctw_xml_elkton",
    "ctw_xml_lb",
    "benediction_text",
    "benediction_text_alt",
    "announce_detail_elkton",
    "announce_detail_lb",
})


def _expand_writer_paragraph(text: str, name: str, value: str) -> str:
    """Replace {name} with multiple <text:p> elements, one per non-empty line.

    Finds the surrounding <text:p ...> opening tag in *text* and replicates it
    for each line of *value*.  Any trailing XML between the placeholder and the
    closing </text:p> (e.g. a stray <text:tab/>) is discarded.
    """
    pattern = r'(<text:p[^>]*>)\{' + re.escape(name) + r'\}.*?(</text:p>)'
    m = re.search(pattern, text, re.DOTALL)
    if not m:
        # Placeholder not found as a standalone paragraph — fall back to simple insert.
        return text.replace("{" + name + "}", value)
    opening_tag = m.group(1)
    closing_tag = m.group(2)
    lines = [ln for ln in value.split("\n") if ln.strip()]
    is_ctw = name.startswith("ctw_xml_")
    if not lines:
        replacement = opening_tag + closing_tag
    else:
        parts = []
        for ln in lines:
            if is_ctw and ln.strip().startswith("P"):
                content = '<text:span text:style-name="Strong Emphasis">' + ln + "</text:span>"
            else:
                content = ln
            parts.append(opening_tag + content + closing_tag)
        replacement = "".join(parts)
    return text[: m.start()] + replacement + text[m.end() :]


def _format_writer_placeholder_value(placeholder_name: str, value: str) -> str:
    """
    Writer/ODT content.xml requires XML-safe values.
    - Escape '&' as '&amp;' for all values (content.xml is XML).
    - Replace newlines with <text:line-break/> for multi-paragraph placeholders.
    """
    value = value.replace("&", "&amp;")
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
    omit_overrides: dict | None = None,
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

    omit_overrides: dict of {placeholder_name: fallback_value}.  When a placeholder's
      value would trigger the omit rule, use the fallback value instead of deleting
      the line.  Example: {"reader": "nobody"} keeps the reader line in PastorLB.md.
    """
    warnings = []
    omit_overrides = omit_overrides or {}

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
            if candidate in lookup:
                value = lookup[candidate]
                found = True
                break

        if not found:
            msg = f"Placeholder {{{name}}} not found in Master.md (left unchanged)."
            if strict:
                raise ValueError(msg)
            warnings.append(msg)
            continue

        value = _apply_delete_rule(value)
        value = _reflow_for_public_reading(name, value)

        # Apply omit before actual replacement (needs the token still present),
        # unless this placeholder has an omit override (use the fallback instead).
        if name in omit_overrides and value.strip().lower() == "omit":
            value = omit_overrides[name]
        else:
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


def _liturgist_output_path(worship_dir: Path, date_slug: str) -> Path:
    return worship_dir / f"Liturgist--{date_slug}.md"




# -----------------------------------------------------------------------------
# LibreOffice Writer output builder (stub)
# -----------------------------------------------------------------------------

def build_librewriter_outputs(*, master_path: Path | None = None, strict: bool = False, verbose: bool = False):
    """
    Access the appropriate subdirectory under bulletin_dir based on the year in {cal_date}.
    This will be expanded to generate Writer files from templates in that subdirectory.
    """
    if master_path is None:
        raise ValueError("master_path must be provided")
    if not master_path.exists():
        raise FileNotFoundError(f"Master.md not found at: {master_path}")

    master_md = read_text(master_path)
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
        m = re.search(r"(20\d{2})", cal_date)
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

    # Determine which segments to run (default: all)
    run_md = not (args.openlp or args.writer)
    run_wr = not (args.openlp or args.markdown)
    run_ol = not (args.markdown or args.writer)

    speaker_results, liturgist_result, writer_results = build_markdown_and_writer_outputs(
        master_path=master_path,
        strict=args.strict,
        verbose=args.verbose,
        run_markdown=run_md,
        run_writer=run_wr,
    )

    for church, result in speaker_results.items():
        print(f"Speaker file for {church}: {result.output_path.name}")
        if result.warnings:
            print(f"  Warnings for {church}:")
            for w in result.warnings:
                print("   -", w)

    if liturgist_result is not None:
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

    if run_ol:
        openlp_results = copy_openlp_templates_for_each_church(
            master_path=master_path,
            strict=args.strict,
            verbose=args.verbose,
        )
        for church, out_path in openlp_results.items():
            print(f"OpenLP file for {church}: {out_path.name}")
