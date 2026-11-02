"""Matching and rewriting items inside an OpenLP service file."""

from __future__ import annotations

import json
import zipfile

import scripts.publish as pub


def titles(service_data: list, plugin: str | None = None) -> list[str]:
    out = []
    for entry in service_data:
        svc = entry.get("serviceitem") if isinstance(entry, dict) else None
        if not isinstance(svc, dict):
            continue
        header = svc.get("header") or {}
        if plugin and str(header.get("plugin", "")).lower() != plugin:
            continue
        out.append(str(header.get("title", "")))
    return out


def find(service_data: list, title: str) -> dict:
    for entry in service_data:
        svc = entry.get("serviceitem") if isinstance(entry, dict) else None
        if isinstance(svc, dict) and str(svc["header"].get("title", "")) == title:
            return svc
    raise AssertionError(f"no service item titled {title!r}")


def test_offertory_matches_suffixed_title(service_data, custom_db):
    """
    The Elkton communion templates title the item 'Offertory Prayer - GBOD'. Exact
    matching skipped it silently and last week's prayer went to the projector.
    """
    data = service_data("svc_elk_communion1.osj")
    assert "Offertory Prayer - GBOD" in titles(data, "custom")

    ok = pub._replace_offertory_prayer_content_and_credits(
        data, "elkton", "This week's prayer. Amen.", "Copyright GBOD."
    )
    assert ok is True

    svc = find(data, "Offertory Prayer - GBOD")
    assert "This week's prayer. Amen." in svc["data"][0]["raw_slide"]
    # The title is half the key OpenLP re-links the library slide by: leave it alone.
    assert svc["header"]["title"] == "Offertory Prayer - GBOD"
    assert svc["header"]["footer"] == ["Offertory Prayer - GBOD Copyright GBOD."]


def test_offertory_matches_bare_title(service_data, custom_db):
    data = service_data("svc_lb.osj")
    assert pub._replace_offertory_prayer_content_and_credits(
        data, "lb", "Prayer body.", "Copyright GBOD."
    )
    assert find(data, "Offertory Prayer")["data"][0]["raw_slide"].endswith("Prayer body.\n")


def test_offertory_empty_master_text_is_not_applied(service_data, custom_db):
    data = service_data("svc_lb.osj")
    before = json.dumps(data)
    assert pub._replace_offertory_prayer_content_and_credits(data, "lb", "", "c") is False
    assert json.dumps(data) == before, "a no-op must leave the template untouched"


def test_benediction_three_slides(service_data, custom_db):
    data = service_data("svc_lb.osj")
    text = "Go now in faith\nand live.\n\nMay God keep you;\nand give you peace."
    assert pub._replace_benediction_custom_slide(data, "lb", text) is True
    slides = find(data, "Benediction")["data"]
    assert [s["verseTag"] for s in slides] == ["1", "2", "3"]
    assert slides[0]["raw_slide"] == "{y}Go now in faith and live.{/y}"
    assert "In the name of Christ. Amen." in slides[2]["raw_slide"]


def test_benediction_single_paragraph_is_rejected(service_data, custom_db):
    """Fewer than two paragraphs cannot fill the three slides, so nothing is written."""
    data = service_data("svc_lb.osj")
    before = json.dumps(data)
    assert pub._replace_benediction_custom_slide(data, "lb", "Only one paragraph.") is False
    assert json.dumps(data) == before


def test_reading_second_line_replacement(service_data):
    data = service_data()
    assert pub._replace_custom_item_second_line(data, "Reading: 1", "Exodus 20:1-4") is True
    raw = find(data, "Reading: 1")["data"][0]["raw_slide"]
    assert raw.splitlines()[1] == "Exodus 20:1-4"


def test_reading_miss_is_reported_by_return_value(service_data):
    """
    publish currently discards this result, so a renamed template item silently keeps
    last week's reference. The matcher itself must at least report the miss.
    """
    data = service_data()
    assert pub._replace_custom_item_second_line(data, "Reading: Nope", "Exodus 20") is False


def test_malformed_entries_are_skipped(service_data):
    """The leading {'openlp_core': ...} entry and junk items must not raise."""
    data = service_data()
    data.insert(0, {"openlp_core": {"lite-service": False}})
    data.append({"serviceitem": {"header": None}})
    data.append({"serviceitem": "not a dict"})
    data.append("not a dict either")
    assert pub._replace_custom_item_second_line(data, "Reading: 1", "Exodus 20:1-4") is True


def test_osz_round_trip_preserves_other_members(osz_factory):
    path = osz_factory()
    data = pub._load_service_data_from_osz(path)
    pub._replace_custom_item_second_line(data, "Reading: 1", "Exodus 20:1-4")
    pub._write_service_data_to_osz(path, data)

    with zipfile.ZipFile(path) as zf:
        assert "deadbeef.avi" in zf.namelist()
        assert zf.read("deadbeef.avi") == b"not-a-real-video"
    reloaded = pub._load_service_data_from_osz(path)
    assert find(reloaded, "Reading: 1")["data"][0]["raw_slide"].splitlines()[1] == "Exodus 20:1-4"
    assert len(reloaded) == len(data), "no items added or dropped by the round trip"
