#!/usr/bin/env python3
from __future__ import annotations
import logging
import re
import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Tuple
from datetime import datetime
import json
import shutil
import zipfile
import tempfile
from xml.etree import ElementTree as ET

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.utils.config import config, require_dir, resolve_master_path
from scripts.utils.file_io import atomic_write, read_text, replace_file_atomic, write_text_atomic
from scripts.utils.logging_utils import init_logging
from scripts.utils.placeholder import (
    extract_block,
    has_placeholder,
    master_value,
    omitted_marker,
)
from scripts.utils.osz import (
    find_service_item,
    find_service_items,
    has_service_item,
    load_service_data,
    remove_service_items,
    write_service_data,
)
from scripts.utils.hymnal import (
    SongRef,
    hymnal_name,
    intro_slide_prefix,
    parse_song_ref,
    read_verse_filter,
)
from scripts.utils.report import InjectionReport, render_summary
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

def _load_service_data_from_osz(osz_path: Path) -> list:
    """Read service_data.osj from an OpenLP .osz file and return the parsed list."""
    return load_service_data(osz_path)


def _write_service_data_to_osz(osz_path: Path, service_data: list) -> None:
    """Write updated service_data.osj back to an OpenLP .osz, preserving other members."""
    write_service_data(osz_path, service_data)


def _get_slide_text_for_prefix(church: str, prefix: str, exact: bool = False) -> tuple[str, list[str], str] | None:
    """
    Find a custom slide whose title starts with prefix and return
    (title, pages, credits), one page per slide the library record holds.
    When exact=True, only a slide whose normalized title equals prefix is accepted.
    """
    from scripts.utils.openlp import get_custom_slide_pages

    return get_custom_slide_pages(church, prefix, exact=exact)


def _replace_custom_item_text(
    service_data: list,
    marker_title: str,
    new_title: str,
    new_pages: list[str] | str,
    *,
    footer: list | None = None,
    credits: str = "",
) -> bool:
    """
    Replace one custom service item (matched by header.title) with new custom-slide text.
    Returns True if an item was replaced.

    new_pages: one entry per projected slide. A library slide with several pages must
               become several service slides — concatenating them yields one overfull
               slide with a literal [===] marker printed in the middle.
    credits: the Credits field from the source slide; used to build the footer
             when no explicit footer list is provided.
    footer: explicit footer list; when None, defaults to OpenLP's own
            "<title> <credits>" footer string.
    """
    from scripts.utils.openlp import openlp_footer

    item = find_service_item(service_data, plugin="custom", title=marker_title)
    if item is None:
        return False

    item.header["title"] = new_title
    if footer is not None:
        item.header["footer"] = footer
    else:
        # OpenLP parses credits back out of footer[0] as
        # footer[0][len(title) + 1:]; any other shape breaks the library
        # link, which disables Edit and makes OpenLP clone the slide.
        item.header["footer"] = [openlp_footer(new_title, credits)]

    pages = [new_pages] if isinstance(new_pages, str) else list(new_pages)
    item.svc["data"] = [
        {"title": new_title[:30], "raw_slide": page, "verseTag": str(idx)}
        for idx, page in enumerate(pages, start=1)
    ]
    return True


def _replace_custom_item_second_line(service_data: list, marker_title: str, new_second_line: str) -> bool:
    """
    For a custom slide matched by header.title, keep line 1 of raw_slide unchanged
    and replace line 2 with new_second_line.
    """
    item = find_service_item(service_data, plugin="custom", title=marker_title)
    if item is None:
        return False

    data = item.svc.get("data")
    if not isinstance(data, list) or not data:
        return False

    first = data[0] if isinstance(data[0], dict) else {}
    lines = str(first.get("raw_slide", "") or "").splitlines()
    if not lines:
        return False

    # Keep line 1, replace line 2, preserve any remaining lines.
    replacement = (new_second_line or "").strip()
    if len(lines) == 1:
        lines.append(replacement)
    else:
        lines[1] = replacement

    first["raw_slide"] = "\n".join(lines)
    data[0] = first
    return True


def _replace_offertory_prayer_content_and_credits(
    service_data: list,
    church: str,
    body_text: str,
    credit_text: str,
) -> bool:
    """
    Update the custom slide whose title starts with 'Offertory Prayer' (templates
    use suffixed titles such as 'Offertory Prayer - GBOD'):
      - Rebuild slide data from scratch to avoid retaining stale template text
      - Replace body text with body_text
      - Update header.footer so credits project from offertory_credit
      - Write the same text into the church's custom slide library, so OpenLP's
        Edit shows this week's prayer instead of the stored one
    """
    from scripts.utils.openlp import credits_from_footer, openlp_footer, sync_custom_slide

    body = (body_text or "").strip()
    if not body:
        return False

    credits_raw = (credit_text or "").replace("\r\n", "\n").replace("\r", "\n")
    credits = " ".join(line.strip() for line in credits_raw.split("\n") if line.strip())

    item = find_service_item(
        service_data, plugin="custom", title="Offertory Prayer", match="prefix"
    )
    if item is None:
        return False

    # Keep the template's title: it is half of the key OpenLP uses to re-link the
    # item to its library slide on load.
    title = item.title
    if not credits:
        credits = credits_from_footer(title, item.header.get("footer"))

    slide_text = f"{{h}}Offertory Prayer{{/h}}\n\n{body}\n"
    item.svc["data"] = [{
        "title": "Offertory Prayer",
        "raw_slide": slide_text,
        "verseTag": "1",
    }]
    item.header["footer"] = [openlp_footer(title, credits)]

    sync_custom_slide(church, title, credits, [slide_text])
    return True


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


def _replace_benediction_custom_slide(service_data: list, church: str, benediction_text: str) -> bool:
    """
    Update custom slide titled 'Benediction':
      - Keep the first raw_slide line unchanged (heading/markup)
      - Replace body with formatted benediction content
      - Write the same slides into the church's custom slide library, so OpenLP's
        Edit shows this week's benediction instead of the stored one
    """
    from scripts.utils.openlp import credits_from_footer, sync_custom_slide

    slides = _format_benediction_slides_for_openlp(benediction_text)
    if not slides:
        return False

    # Prefix match: a template that retitles this item "Benediction - Advent" must
    # not silently keep last week's words, which is what exact matching allowed.
    item = find_service_item(
        service_data, plugin="custom", title="Benediction", match="prefix"
    )
    if item is None:
        return False

    title = item.title
    item.svc["data"] = [
        {"title": title[:30], "raw_slide": slide, "verseTag": str(idx)}
        for idx, slide in enumerate(slides, start=1)
    ]
    sync_custom_slide(
        church,
        title,
        credits_from_footer(title, item.header.get("footer")),
        slides,
    )
    return True


def _remove_service_items_by_marker(service_data: list, plugin: str, marker_title: str) -> int:
    """Remove all service items matching plugin + header.title (case-insensitive). Returns count removed."""
    if not (plugin or "").strip() or not (marker_title or "").strip():
        return 0
    return remove_service_items(service_data, plugin=plugin, title=marker_title)


def _build_openlyrics_xml(
    title: str,
    authors: str,
    copyright_text: str,
    hymnal: str,
    entry: str,
    slides: list[dict],
) -> str:
    """Build a minimal OpenLyrics 0.8 XML string for header.xml_version."""
    from datetime import datetime

    ns = "http://openlyrics.info/namespace/2009/song"
    ET.register_namespace("", ns)

    song_el = ET.Element(f"{{{ns}}}song")
    song_el.set("version", "0.8")
    song_el.set("createdIn", "OpenLP 3.1.7")
    song_el.set("modifiedIn", "OpenLP 3.1.7")
    song_el.set("modifiedDate", datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))

    props = ET.SubElement(song_el, f"{{{ns}}}properties")
    titles_el = ET.SubElement(props, f"{{{ns}}}titles")
    ET.SubElement(titles_el, f"{{{ns}}}title").text = title or ""

    if copyright_text:
        ET.SubElement(props, f"{{{ns}}}copyright").text = copyright_text

    if authors:
        authors_el = ET.SubElement(props, f"{{{ns}}}authors")
        for author in [a.strip() for a in authors.split(",") if a.strip()]:
            ET.SubElement(authors_el, f"{{{ns}}}author").text = author

    if hymnal and entry:
        books = ET.SubElement(props, f"{{{ns}}}songbooks")
        sb = ET.SubElement(books, f"{{{ns}}}songbook")
        sb.set("name", hymnal)
        sb.set("entry", entry)

    lyrics_el = ET.SubElement(song_el, f"{{{ns}}}lyrics")
    seen: set[str] = set()
    for slide in (slides or []):
        tag = str(slide.get("verseTag") or "")
        if not tag:
            continue
        name = tag.lower()  # V1→v1, C1→c1
        if name in seen:
            continue
        seen.add(name)
        verse_el = ET.SubElement(lyrics_el, f"{{{ns}}}verse")
        verse_el.set("name", name)
        ET.SubElement(verse_el, f"{{{ns}}}lines").text = str(slide.get("raw_slide") or "")

    return "<?xml version='1.0' encoding='UTF-8'?>\n" + ET.tostring(song_el, encoding="unicode")


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

    item = find_service_item(service_data, plugin="songs", title=marker_title)
    if item is not None:
        svc, header = item.svc, item.header
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
        header["xml_version"] = _build_openlyrics_xml(
            new_title, authors_clean, copyright_clean, hymnal_clean, entry_clean, slides
        )
        return True

    return False


def _fetch_scripture_verses(church: str, passage: str) -> list[tuple[int, int, str]]:
    """
    Scripture verses as (chapter, verse, text).

    Thin wrapper over the one parser in utils.openlp — publish used to carry a
    near-identical ~100-line copy, and the two had drifted over how a malformed
    chapter number is handled.
    """
    from scripts.utils.openlp import fetch_scripture_verses

    return list(fetch_scripture_verses(church, passage))


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
    bible_items = find_service_items(service_data, plugin="bibles")
    if len(bible_items) < bible_order:
        return False

    svc = bible_items[bible_order - 1].svc
    header = bible_items[bible_order - 1].header

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


def _has_service_item(service_data: list, plugin: str, title: str) -> bool:
    """True when the template carries an item with this plugin and title."""
    return has_service_item(service_data, plugin=plugin, title=title)


def _inject_custom_slides_into_openlp_service(
    osz_path: Path,
    church: str,
    master_md: str,
    *,
    communion: str = "",
    template_name: str = "",
    report: InjectionReport | None = None,
) -> InjectionReport:
    """
    Fill a copied OpenLP service archive from the master: Call to Worship, Affirmation
    of Faith, scripture, songs, readings, offertory prayer and benediction.

    Returns an InjectionReport. Every step records a finding, so a step that matched
    nothing is visible in the run summary instead of quietly leaving the template's
    own content in the published service.
    """
    from scripts.utils.openlp import get_scripture_text

    ctw_ref = master_value(master_md, "ctw_ref")
    ctw_ref_church = master_value(master_md, f"ctw_ref_{church}")
    aof_ref = master_value(master_md, "aof_ref")
    scripture_ref_1 = master_value(master_md, "scripture_ref_1")
    scripture_ref_2 = master_value(master_md, "scripture_ref_2")
    scripture_ref_3 = master_value(master_md, "scripture_ref_3")
    offertory_source_text = master_value(master_md, "offertory_source_text")
    offertory_credit = master_value(master_md, "offertory_credit")
    benediction_text = master_value(master_md, "benediction_text")

    # Songs (OpenLP markers live in the service templates)
    song_opening_title = master_value(master_md, f"song_opening_title_{church}")
    song_opening_id = master_value(master_md, f"song_opening_id_{church}")
    song_middle_title = master_value(master_md, f"song_middle_title_{church}")
    song_middle_id = master_value(master_md, f"song_middle_id_{church}")
    song_closing_title = master_value(master_md, f"song_closing_title_{church}")
    song_closing_id = master_value(master_md, f"song_closing_id_{church}")

    service_data = _load_service_data_from_osz(osz_path)
    report = report if report is not None else InjectionReport(church=church, communion=communion)

    if not ctw_ref_church:
        report.skip("ctw", "no ctw_ref set in the master")
    else:
        ctw = _get_slide_text_for_prefix(church, ctw_ref_church, exact=True)
        if not ctw:
            report.error("ctw", f"no library slide titled {ctw_ref_church!r}")
        else:
            ctw_title, ctw_pages, ctw_credits = ctw
            if _replace_custom_item_text(service_data, "ctw_holder", ctw_title, ctw_pages, credits=ctw_credits):
                report.ok("ctw", f"{ctw_title} ({len(ctw_pages)} slide(s))")
            else:
                report.error("ctw", "no 'ctw_holder' item in the service template")

    if not aof_ref:
        report.skip("aof", "no aof_ref set in the master")
    elif not _has_service_item(service_data, "custom", "aof_holder"):
        # aof_holder exists only in the non-communion Elkton template, so its
        # absence on a communion Sunday is normal, not a failure.
        report.skip("aof", f"no 'aof_holder' item in {template_name or 'this template'}")
    else:
        aof = _get_slide_text_for_prefix(church, f"AoF p{aof_ref}")
        if not aof:
            report.error("aof", f"no library slide matching 'AoF p{aof_ref}'")
        else:
            aof_title, aof_pages, aof_credits = aof
            if _replace_custom_item_text(service_data, "aof_holder", aof_title, aof_pages, credits=aof_credits):
                report.ok("aof", f"{aof_title} ({len(aof_pages)} slide(s))")
            else:
                report.error("aof", "no 'aof_holder' item in the service template")

    # Scripture updates:
    # 1) Prefer native OpenLP Bible plugin items by order (1, 2, 3)
    # 2) Fallback to custom holder slides if no bibles items are found
    bible_item_count = len(find_service_items(service_data, plugin="bibles"))

    refs = ((1, scripture_ref_1), (2, scripture_ref_2), (3, scripture_ref_3))
    if bible_item_count >= 3:
        for idx, ref in refs:
            slot = f"bible_{idx}"
            if not ref:
                report.skip(slot, f"no scripture_ref_{idx} set in the master")
                continue
            try:
                if _replace_bible_item_by_order(service_data, idx, ref, church):
                    report.ok(slot, ref)
                else:
                    report.error(slot, f"Bible service item #{idx} not found in the template")
            except Exception as e:
                report.error(slot, f"{ref}: {type(e).__name__}: {e}")
    else:
        # No bibles items: every shipped template has three, so this only happens
        # if a template is rebuilt without them.
        for idx, ref in refs:
            report.error(
                f"bible_{idx}",
                f"template has {bible_item_count} Bible item(s); expected 3",
            )

    # Song updates (marker-based)
    try:
        _inject_songs_into_openlp_service_data(
            service_data,
            church,
            opening_title=song_opening_title,
            opening_id=song_opening_id,
            middle_title=song_middle_title,
            middle_id=song_middle_id,
            closing_title=song_closing_title,
            closing_id=song_closing_id,
            report=report,
            verse_filters={
                slot: read_verse_filter(master_md, slot, church)
                for slot in ("opening", "middle", "closing")
            },
        )
    except Exception as e:
        report.error("songs", f"{type(e).__name__}: {e}")

    # Reading custom slides: keep first line, replace second line with scripture reference
    reading_map = (
        ("reading_1", "Reading: 1", scripture_ref_1),
        ("reading_2", "Reading: 2", scripture_ref_2),
        ("reading_gospel", "Reading: Gospel", scripture_ref_3),
    )
    for slot, reading_title, reading_ref in reading_map:
        if not reading_ref:
            report.skip(slot, "no scripture reference set in the master")
            continue
        if _replace_custom_item_second_line(service_data, reading_title, reading_ref.strip()):
            report.ok(slot, reading_ref)
        else:
            report.error(slot, f"no custom item titled {reading_title!r}")

    if not offertory_source_text:
        report.skip("offertory_prayer", "no offertory_source_text set in the master")
    elif _replace_offertory_prayer_content_and_credits(
        service_data,
        church,
        offertory_source_text,
        offertory_credit,
    ):
        report.ok("offertory_prayer")
    else:
        report.error("offertory_prayer", "no custom item whose title starts with 'Offertory Prayer'")

    if not benediction_text:
        report.skip("benediction", "no benediction_text set in the master")
    elif _replace_benediction_custom_slide(service_data, church, benediction_text):
        report.ok("benediction")
    elif _format_benediction_slides_for_openlp(benediction_text) is None:
        report.error("benediction", "benediction_text needs two blank-line separated paragraphs")
    else:
        report.error("benediction", "no custom item titled 'Benediction'")

    if report.changed:
        _write_service_data_to_osz(osz_path, service_data)
    else:
        report.stale = True

    return report


def _inject_songs_into_openlp_service_data(service_data: list, church: str, *,
                                          opening_title: str, opening_id: str,
                                          middle_title: str, middle_id: str,
                                          closing_title: str, closing_id: str,
                                          report: InjectionReport,
                                          verse_filters: dict[str, tuple[int, ...] | None] | None = None) -> bool:
    """Return True if any song items were updated/removed using ID-driven slot rules."""
    from scripts.utils.text_clean import xml_to_text
    from scripts.utils.openlp import build_song_index, load_song

    changed = False
    church_key = (church or "").strip().lower()
    verse_filters = verse_filters or {}

    def warn_song_missing(slot_marker: str, message: str) -> None:
        report.error(slot_marker, message)

    def apply_intro_from_custom_slide(holder_marker: str, title: str, ref: SongRef | None) -> bool:
        intro_prefix = intro_slide_prefix(church_key, ref, title)
        if not intro_prefix:
            return False
        slide = _get_slide_text_for_prefix(church_key, intro_prefix)
        if not slide:
            report.warn(holder_marker, f"no intro library slide matching {intro_prefix!r}")
            return False
        slide_title, slide_pages, slide_credits = slide
        return _replace_custom_item_text(service_data, holder_marker, slide_title, slide_pages, credits=slide_credits)

    song_index = None

    def try_replace_song(song_marker: str, title: str, ref: SongRef, verses: tuple[int, ...] | None) -> bool:
        nonlocal song_index
        prefix, entry = ref.prefix, ref.entry
        hymnal = hymnal_name(church_key, prefix)
        if not hymnal:
            warn_song_missing(song_marker, f"unknown hymnal prefix {prefix!r} for {church_key}")
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
        pre_slides = parse_song_xml(lyrics_xml, verse_order, verses_filter=verses)
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
        slot_id = (str(slot_id_raw or "").split("%%", 1)[0]).strip()

        if not slot_id:
            report.skip(song_marker, "no song id set in the master")
            continue

        # tbd means remove both intro + song placeholders for this slot
        if slot_id.lower() == "tbd":
            removed_custom = _remove_service_items_by_marker(service_data, "custom", holder_marker)
            removed_song = _remove_service_items_by_marker(service_data, "songs", song_marker)
            if removed_custom or removed_song:
                changed = True
            report.skip(song_marker, "id is 'tbd' - intro and song items removed")
            continue

        ref = parse_song_ref(slot_id, title_for_log=slot_title)
        if ref is None:
            # Video/other non-hymnal ID: pick intro custom slide by P901 <title>, leave song placeholder unchanged.
            if apply_intro_from_custom_slide(holder_marker, slot_title, None):
                changed = True
            report.skip(song_marker, f"id {slot_id!r} is not a hymnal reference - media slot left for manual insert")
            continue

        verses = verse_filters.get(slot_name)
        if ref.legacy_verses and verses is None:
            report.warn(
                song_marker,
                f"verse list in the song id is ignored - put it in "
                f"{{song_{slot_name}_vv_{church_key}}}",
            )

        # For hymnal IDs, intro slide selection is based on prefix+number convention.
        if apply_intro_from_custom_slide(holder_marker, slot_title, ref):
            changed = True

        # Replace actual song item when resolvable; if not found, leave placeholder item unchanged.
        if try_replace_song(song_marker, slot_title, ref, verses):
            changed = True
            detail = f"{slot_id} {slot_title}".strip()
            if verses:
                detail += f" (verses {', '.join(str(v) for v in verses)})"
            report.ok(song_marker, detail)

    return changed


def _copy_osz_to_tmp(src_path: Path) -> "Path | None":
    """Copy src_path to a validated local /tmp file and return that path.

    Reads from the OneDrive FUSE mount in 1MB chunks to avoid the 10MB
    sequential-read boundary issue, validates the full zip structure (central
    directory, not just EOCD), and returns the /tmp Path on success.  Retries
    once.  Returns None if both attempts fail.
    """
    chunk_size = 1024 * 1024  # 1MB

    for attempt in range(1, 3):
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".osz") as tmp:
                tmp_path = Path(tmp.name)

            with src_path.open("rb") as src, tmp_path.open("wb") as dst:
                shutil.copyfileobj(src, dst, chunk_size)

            with zipfile.ZipFile(tmp_path, "r") as z:
                z.namelist()
                z.getinfo("service_data.osj")

            return tmp_path

        except Exception:
            if tmp_path:
                tmp_path.unlink(missing_ok=True)
            if attempt == 1:
                logging.warning(
                    "Copy of %s to /tmp is not a valid zip (attempt %d/2), retrying...",
                    src_path.name, attempt,
                )

    logging.warning(
        "Copy of %s to /tmp still invalid after 2 attempts; skipping.",
        src_path.name,
    )
    return None


def _write_tmp_to_fuse(tmp_path: Path, dst_path: Path) -> None:
    """
    Write a local /tmp file to a FUSE mount path in 1MB chunks, via a temp file
    next to the destination, so a failure part way through cannot destroy the
    service file that is already there.
    """
    replace_file_atomic(tmp_path, dst_path)


def copy_openlp_templates_for_each_church(
    *,
    master_path: Path | None = None,
    verbose: bool = False,
    dry_run: bool = False,
    reports: Dict[str, InjectionReport] | None = None,
) -> Dict[str, Path]:
    """
    For each church, copy the OpenLP template file to a new file in the source directory
    with a date-based name, using config and file I/O utilities.
    Returns a dict mapping church name to the new file path.

    Each church's findings are collected into `reports` (keyed by church) so the caller
    can print one summary; with dry_run nothing is written to the OpenLP folders.
    """
    init_logging(verbose=verbose)

    if reports is None:
        reports = {}

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
            reports[church] = InjectionReport(church=church, communion=communion)
            reports[church].error("template", f"OpenLP root not found: {openlp_root}")
            continue
        if not template_name:
            reports[church] = InjectionReport(church=church, communion=communion)
            reports[church].error("template", f"no OpenLP template for communion {communion}")
            continue
        # OpenLP .osz files are in 'OpenLP Services' subdir
        service_dir = openlp_root / "OpenLP Services"
        src_path = service_dir / template_name
        if not src_path.exists():
            reports[church] = InjectionReport(church=church, template=template_name, communion=communion)
            reports[church].error("template", f"template not found: {src_path}")
            continue
        # Output to the same subdirectory with 'Service--{date_slug}.osz' as the name
        out_name = f"Service--{date_slug}.osz"
        out_path = service_dir / out_name

        # Work entirely in /tmp to avoid OneDrive FUSE 10MB sequential-read boundary.
        # The file is only written to the FUSE mount once, at the very end, in chunks.
        tmp_copy = _copy_osz_to_tmp(src_path)
        if tmp_copy is None:
            reports[church] = InjectionReport(church=church, template=template_name, communion=communion)
            reports[church].error("template", f"could not read a valid local copy of {template_name}")
            continue

        report = InjectionReport(church=church, template=template_name, communion=communion)
        reports[church] = report

        try:
            _inject_custom_slides_into_openlp_service(
                tmp_copy, church, master_md,
                communion=communion, template_name=template_name, report=report,
            )
        except Exception as e:
            # The file is still written, but it now holds template content, and the
            # summary has to say so rather than leaving a warning in the log.
            report.error("injection", f"{type(e).__name__}: {e}")
            report.stale = True

        if dry_run:
            logging.info("Dry run: not writing %s", out_path)
            tmp_copy.unlink(missing_ok=True)
            continue

        try:
            _write_tmp_to_fuse(tmp_copy, out_path)
        except Exception as e:
            report.error("write", f"could not write {out_path.name}: {type(e).__name__}: {e}")
            continue
        finally:
            tmp_copy.unlink(missing_ok=True)

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

    bulletin_dir = None
    if run_writer:
        bulletin_dir = require_dir(config.bulletin_dir, "bulletin_dir")

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
            write_text_atomic(md_out, md_rendered)
            speaker_results[church] = RenderResult(md_out, md_warnings)

        if run_writer:
            assert bulletin_dir is not None
            writer_template_name = templates[f"{church}_writer"]
            writer_template_path = bulletin_dir / writer_template_name
            if not writer_template_path.exists():
                logging.warning(f"Writer template not found: {writer_template_path}")
                writer_results[church] = None
            else:
                writer_out = bulletin_dir / f"{date_slug}-{Path(writer_template_name).stem}.odt"
                with zipfile.ZipFile(writer_template_path, 'r') as zin, \
                        atomic_write(writer_out, mode="wb") as odt_handle, \
                        zipfile.ZipFile(odt_handle, 'w') as zout:
                        for item in zin.infolist():
                            data = zin.read(item.filename)
                            if item.filename == 'content.xml':
                                text = data.decode('utf-8')
                                placeholders = extract_placeholders_used(text)
                                lookup = build_master_lookup(master_md, placeholders)
                                for name in placeholders:
                                    if name not in lookup:
                                        text = text.replace(f"{{{name}}}", "")
                                        continue
                                    value = lookup[name]
                                    # In XML, "omit" and "delete" both collapse to empty string
                                    # (line-deletion would destroy ODT XML structure)
                                    if omitted_marker(value):
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
        write_text_atomic(liturgist_out, liturgist_rendered)
        liturgist_result = RenderResult(liturgist_out, liturgist_warnings)
        for w in liturgist_warnings:
            logging.warning("[Liturgist] %s", w)

    for church, result in speaker_results.items():
        for w in result.warnings:
            logging.warning(f"[Speaker {church}] %s", w)

    return speaker_results, liturgist_result, writer_results


def parse_args(argv: list[str] | None = None):
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
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report what the OpenLP injection would do without writing any service file",
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
    return parser.parse_args(argv)



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

        if re.match(r"^#{1,2}\s+\S", stripped):
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


# Liturgical cue labels that each begin a new Call to Worship paragraph.
_CTW_CUE_RE = re.compile(r"^(?:Leader|People|All|L|P)\s*:", re.IGNORECASE)


def _collapse_ctw_to_cue_lines(value: str) -> str:
    """
    One line per liturgical cue: a line starting with Leader:/L:/People:/P:/All:
    begins a new paragraph, and any hard-wrapped continuation lines are folded
    into it.

    Collapsing on the master's blank lines instead (the old behaviour) made
    paragraph breaks depend on how the source happened to be wrapped, so two
    cues sharing a block ran together in the bulletin.
    """
    paragraphs: list[str] = []
    for raw_line in value.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = re.sub(r"\s+", " ", raw_line.strip())
        if not line:
            continue
        if _CTW_CUE_RE.match(line) or not paragraphs:
            paragraphs.append(line)
        else:
            paragraphs[-1] = f"{paragraphs[-1]} {line}"
    return "\n".join(paragraphs)


def _preserve_line_structure(value: str) -> str:
    """Keep the author's own line breaks, trimming trailing space and blank runs."""
    normalized = value.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in normalized.split("\n")]
    return "\n".join(line for line in lines if line.strip())


def _format_multiline_fields_for_templates(placeholder_name: str, value: str) -> str:
    """
    For selected placeholders, keep one newline between paragraphs while
    collapsing wrapped lines inside each paragraph to spaces.

    Two exceptions:
      - announce_event(s)_*: each event stays on its own line, as written in the master.
      - ctw_xml_*: one line per liturgical cue (see _collapse_ctw_to_cue_lines).
    """
    if not _requires_single_newline_paragraphs(placeholder_name):
        return value

    if placeholder_name.startswith(("announce_event_", "announce_events_")):
        return _preserve_line_structure(value)

    if placeholder_name.startswith("ctw_xml_"):
        return _collapse_ctw_to_cue_lines(value)

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


# Markdown emphasis per Call to Worship speaker, keyed by the cue's first letter.
_CTW_MARKDOWN_STYLE = {"l": "**", "a": "**", "p": "*"}


def _format_ctw_markdown_lines(name: str, value: str) -> str:
    """
    For ctw_xml_* placeholders, style each speaker's part: Leader:/L:/All: in
    bold, People:/P: in italics. A cue (the same set the bulletin recognises,
    _CTW_CUE_RE) sets the style for its own line and the continuation lines under
    it, up to the next cue; lines before the first cue stay plain.

    Each line is wrapped on its own because Markdown emphasis cannot span a blank
    line. Matching any line that merely started with L or P instead styled
    ordinary text such as "Let them praise…" and left the markers unclosed.
    """
    if not name.startswith("ctw_xml_"):
        return value
    result = []
    style = ""
    for line in value.split("\n"):
        stripped = line.strip()
        if _CTW_CUE_RE.match(stripped):
            style = _CTW_MARKDOWN_STYLE[stripped[0].lower()]
        if stripped and style:
            indent = line[: len(line) - len(line.lstrip())]
            line = f"{indent}{style}{stripped}{style}"
        result.append(line)
    return "\n".join(result)


# Bulletin Call to Worship lines set in Strong Emphasis: the congregation's parts.
_CTW_BULLETIN_EMPHASIS_RE = re.compile(r"^(?:People|All|P)\s*:", re.IGNORECASE)


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
            if is_ctw and _CTW_BULLETIN_EMPHASIS_RE.match(ln.strip()):
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


def _apply_delete_rule(value: str) -> str:
    """
    If Master contains 'delete', replace with empty string.
    This mirrors your Master.md conventions.
    """
    return "" if value.strip().lower() == "delete" else value


def _omit_discard_warning(name: str, value: str, marker: str) -> str | None:
    """Warn when text under an omit marker is being dropped, so it is never silent."""
    remainder = "\n".join(
        line for line in value.split("\n")[1:] if line.strip()
    ).strip()
    if not remainder:
        return None
    preview = re.sub(r"\s+", " ", remainder)[:60]
    return f"Placeholder {{{name}}} is '{marker}'; discarded the text under it: {preview!r}"


def _drop_lines_for_omitted_placeholder(template_text: str, name: str) -> str:
    """
    Remove template lines whose only content is this placeholder, so an omitted
    value takes its label line with it instead of leaving a blank or a dangling
    label. Lines holding other text keep their position and lose only the token.
    """
    token = f"{{{name}}}"
    kept = []
    for line in template_text.split("\n"):
        if token in line and not line.replace(token, "").strip():
            continue
        kept.append(line)
    return "\n".join(kept)


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
      - omit/delete rule: a block whose first non-empty line is 'omit' or 'delete'
        renders as nothing, and takes its template line with it when the placeholder
        stands alone on that line. Text found under the marker is discarded, with a
        warning naming what was dropped.

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
            msg = f"Placeholder {{{name}}} not found in Master.md (removed)."
            if strict:
                raise ValueError(msg)
            warnings.append(msg)
            rendered = rendered.replace(f"{{{name}}}", "")
            continue

        marker = omitted_marker(value)
        if marker and name in omit_overrides:
            value = omit_overrides[name]
        elif marker:
            discard = _omit_discard_warning(name, value, marker)
            if discard:
                warnings.append(discard)
            # Take the whole line with it where the placeholder stands alone.
            rendered = _drop_lines_for_omitted_placeholder(rendered, name)
            value = ""
        else:
            value = _apply_delete_rule(value)

        value = _reflow_for_public_reading(name, value)
        value = _format_ctw_markdown_lines(name, value)

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
    """Return cal_date as YYYY-MM-DD for filenames; fall back to sanitized raw text."""
    from datetime import datetime
    raw = extract_block(master_md, "cal_date").strip()
    raw = raw or "unknown-date"
    raw = clean_text(raw)
    for fmt in ("%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    slug = re.sub(r"[^\w\-]+", "-", raw)
    slug = re.sub(r"-{2,}", "-", slug).strip("-")
    return slug.lower()


def _liturgist_output_path(worship_dir: Path, date_slug: str) -> Path:
    return worship_dir / f"Liturgist--{date_slug}.md"




# -----------------------------------------------------------------------------
# LibreOffice Writer output builder (stub)
# -----------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    """Run the publish steps and return a process exit code."""
    args = parse_args(argv)

    master_path = resolve_master_path(args.master)

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

    reports: Dict[str, InjectionReport] = {}
    if run_ol:
        openlp_results = copy_openlp_templates_for_each_church(
            master_path=master_path,
            verbose=args.verbose,
            dry_run=args.dry_run,
            reports=reports,
        )
        destinations = {church: path.name for church, path in openlp_results.items()}
        if args.dry_run:
            destinations = {church: "(dry run, not written)" for church in reports}
        print(render_summary(reports, destinations))

    # Files are always written; --strict turns an incomplete run into a failing
    # exit code so a wrapper can notice.
    failed = any(r.errors or not r.changed for r in reports.values())
    return 1 if (failed and args.strict) else 0


if __name__ == "__main__":
    sys.exit(main())
