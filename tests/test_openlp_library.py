"""
The link between a service item and the custom-slide library.

OpenLP re-links a custom item on load by title + credits, parsing credits out of
footer[0] as footer[0][len(title) + 1:]. Get that wrong and right-click → Edit is
disabled and OpenLP clones the slide into custom.sqlite on every single load.
"""

from __future__ import annotations

import pytest

from scripts.utils.openlp import (
    build_custom_slide_xml,
    credits_from_footer,
    extract_openlp_slide_pages,
    extract_openlp_slide_text,
    openlp_footer,
    sync_custom_slide,
)


def openlp_parses_credits(title: str, footer: list[str]) -> str:
    """OpenLP's own expression, from plugins/custom/lib/mediaitem.py service_load()."""
    return footer[0][len(title) + 1 :]


@pytest.mark.parametrize(
    "title, credits",
    [
        ("Offertory Prayer - GBOD", "Copyright General Board of Discipleship."),
        ("CtW Psalm 19:7-14", "NRSV"),
        ("Benediction", "©2001 Nathan Nettleton www.laughingbird.net"),
        ("CtW Psalm 19", ""),
    ],
)
def test_footer_round_trips_through_openlps_own_parse(title, credits):
    footer = [openlp_footer(title, credits)]
    assert credits_from_footer(title, footer) == credits
    assert openlp_parses_credits(title, footer).strip() == credits


def test_no_credits_leaves_the_projected_footer_empty():
    """"<title> " would print the title under every slide; "" parses the same."""
    assert openlp_footer("CtW Psalm 19", "") == ""


def test_slide_xml_matches_openlps_shape():
    xml = build_custom_slide_xml(["first slide", "second slide"])
    assert xml.startswith("<?xml version='1.0' encoding='utf-8'?>\n<song version=\"1.0\">")
    assert '<verse type="custom" label="1"><![CDATA[first slide]]></verse>' in xml
    assert '<verse type="custom" label="2"><![CDATA[second slide]]></verse>' in xml
    assert extract_openlp_slide_text(xml) == "first slide\n[===]\nsecond slide"


def test_slide_xml_survives_a_cdata_terminator():
    xml = build_custom_slide_xml(["a ]]> b"])
    assert "]]]]><![CDATA[>" in xml
    assert extract_openlp_slide_text(xml).replace("\n", "") in ("a ]]> b", "a ]][===]> b")


def test_sync_updates_the_matching_row(custom_db):
    row_id = custom_db.add("elkton", "Offertory Prayer - GBOD", "Copyright GBOD.", "<old/>")
    out = sync_custom_slide("elkton", "Offertory Prayer - GBOD", "Copyright GBOD.", ["new body"])
    assert out == row_id
    rows = custom_db.rows("elkton")
    assert len(rows) == 1, "an update must not add a row"
    assert "new body" in rows[0][3]


def test_sync_inserts_when_no_row_matches(custom_db):
    custom_db.add("elkton", "Offertory Prayer", "Different credits", "<old/>")
    out = sync_custom_slide("elkton", "Offertory Prayer", "Copyright GBOD.", ["body"])
    rows = custom_db.rows("elkton")
    assert len(rows) == 2
    assert out == rows[-1][0]
    assert rows[-1][2] == "Copyright GBOD."


def test_sync_keeps_duplicate_rows_in_step(custom_db):
    """Until the duplicates are pruned, OpenLP may link to either one."""
    a = custom_db.add("lb", "Offertory Prayer", "c", "<old/>")
    b = custom_db.add("lb", "Offertory Prayer", "c", "<older/>")
    assert sync_custom_slide("lb", "Offertory Prayer", "c", ["body"]) == a
    texts = {row[0]: row[3] for row in custom_db.rows("lb")}
    assert "body" in texts[a] and "body" in texts[b]


def test_sync_ignores_empty_input(custom_db):
    assert sync_custom_slide("elkton", "Offertory Prayer", "c", []) is None
    assert sync_custom_slide("elkton", "", "c", ["body"]) is None
    assert custom_db.rows("elkton") == []


def test_multi_page_library_slide_becomes_multiple_service_slides(service_data):
    """
    A CtW slide with two pages must produce two service slides. Joining them put a
    literal [===] on the projector in the middle of one overfull slide — which is
    what this week's Elkton service file actually contained.
    """
    import scripts.publish as pub

    pages = extract_openlp_slide_pages(build_custom_slide_xml(["page one", "page two"]))
    assert pages == ["page one", "page two"]

    data = service_data()
    assert pub._replace_custom_item_text(data, "ctw_holder", "CtW Psalm 19", pages, credits="")

    svc = next(
        entry["serviceitem"]
        for entry in data
        if isinstance(entry, dict)
        and isinstance(entry.get("serviceitem"), dict)
        and entry["serviceitem"]["header"].get("title") == "CtW Psalm 19"
    )
    assert [s["raw_slide"] for s in svc["data"]] == ["page one", "page two"]
    assert [s["verseTag"] for s in svc["data"]] == ["1", "2"]
    assert not any("[===]" in s["raw_slide"] for s in svc["data"])
