"""How master values are shaped before they land in a template."""

from __future__ import annotations

import pytest

import scripts.publish as pub


# --- announcements: one event per line ---------------------------------------

def test_events_keep_their_own_lines(master_md):
    value = pub.build_master_lookup(master_md, ["announce_events_elkton"])["announce_events_elkton"]
    assert value.splitlines() == [
        "Today after worship - meeting",
        "Oct 7th 7:00 PM - UPW",
    ]


def test_events_drop_blank_runs_but_not_line_structure():
    value = pub._format_multiline_fields_for_templates(
        "announce_events_elkton", "One  \n\n\nTwo\nThree\n"
    )
    assert value.splitlines() == ["One", "Two", "Three"]


# --- call to worship: a paragraph per cue ------------------------------------

def test_each_cue_starts_a_paragraph(master_md):
    value = pub.build_master_lookup(master_md, ["ctw_xml_elkton"])["ctw_xml_elkton"]
    lines = value.splitlines()
    assert len(lines) == 4
    assert lines[0].startswith("Leader:")
    assert lines[1].startswith("People:")
    assert lines[2].startswith("L:")
    assert lines[3].startswith("P:")


def test_wrapped_continuation_folds_into_its_cue(master_md):
    value = pub.build_master_lookup(master_md, ["ctw_xml_elkton"])["ctw_xml_elkton"]
    assert value.splitlines()[0] == (
        "Leader: The heavens are telling the glory of God; "
        "and the firmament proclaims his handiwork."
    )


@pytest.mark.parametrize("cue", ["Leader:", "People:", "All:", "L:", "P:", "l:", "people:"])
def test_recognised_cue_labels(cue):
    value = pub._collapse_ctw_to_cue_lines(f"First line.\n{cue} second line.")
    assert value.splitlines() == ["First line.", f"{cue} second line."]


def test_text_before_the_first_cue_is_kept():
    value = pub._collapse_ctw_to_cue_lines("A rubric line\nL: a cue")
    assert value.splitlines() == ["A rubric line", "L: a cue"]


def test_markdown_ctw_styles_each_speakers_part():
    value = pub._format_ctw_markdown_lines(
        "ctw_xml_elkton",
        "A rubric\nLeader: The LORD is king;\nlet the peoples tremble!\n\n"
        "Let them praise your name.\nPeople: Holy is he!\nL: Extol\nAll: Amen.",
    )
    assert value.splitlines() == [
        "A rubric",
        "**Leader: The LORD is king;**",
        "**let the peoples tremble!**",
        "",
        "**Let them praise your name.**",
        "*People: Holy is he!*",
        "**L: Extol**",
        "**All: Amen.**",
    ]


def test_markdown_ctw_ignores_other_placeholders():
    assert pub._format_ctw_markdown_lines("aof_xml", "Leader: x") == "Leader: x"


# --- omit / delete ------------------------------------------------------------

@pytest.mark.parametrize(
    "value, expected",
    [
        ("omit", "omit"),
        ("OMIT", "omit"),
        ("delete", "delete"),
        ("omit  \n\nThanks", "omit"),       # marker plus leftover notes
        ("\n\nomit\n", "omit"),
        ("Thanks\nomit", None),             # only the first line counts
        ("omitted entirely", None),
        ("", None),
        ("tbd", None),
    ],
)
def test_omit_marker_detection(value, expected):
    assert pub.omitted_marker(value) == expected


def test_master_value_resolves_omit_to_empty(master_md):
    assert pub.master_value(master_md, "aof_ref") == ""
    assert pub.master_value(master_md, "announce_events_lb") == ""
    assert pub.master_value(master_md, "announce_events_elkton").startswith("Today after worship")


def test_omitted_placeholder_takes_its_line_with_it(master_md):
    template = "Welcome\n\n{announce_events_lb}\n\nAre there any others?\n"
    rendered, warnings = pub.render_markdown_template(template, master_md)
    assert "omit" not in rendered
    assert "Thanks" not in rendered
    assert "Welcome" in rendered and "Are there any others?" in rendered
    assert any("discarded the text under it" in w for w in warnings)


def test_discarded_text_is_reported_not_silent(master_md):
    _, warnings = pub.render_markdown_template("{announce_events_lb}\n", master_md)
    assert [w for w in warnings if "'Thanks'" in w], warnings


def test_omit_override_keeps_the_line(master_md):
    template = "Reader: {announce_events_lb}\n"
    rendered, _ = pub.render_markdown_template(
        template, master_md, omit_overrides={"announce_events_lb": "nobody"}
    )
    assert "Reader: nobody" in rendered


def test_placeholder_sharing_a_line_only_loses_its_token(master_md):
    template = "Events: {announce_events_lb} (end)\n"
    rendered, _ = pub.render_markdown_template(template, master_md)
    assert "Events:" in rendered and "(end)" in rendered


def test_bulletin_ctw_emphasises_people_and_all_cues():
    xml = '<text:p text:style-name="X">{ctw_xml_lb}</text:p>'
    out = pub._expand_writer_paragraph(
        xml, "ctw_xml_lb", "Praise opens this.\nL: One\nP: Two\nPeople: Three\nAll: Four"
    )
    strong = '<text:span text:style-name="Strong Emphasis">'
    assert f"{strong}Praise" not in out
    assert f"{strong}L: One" not in out
    for line in ("P: Two", "People: Three", "All: Four"):
        assert f"{strong}{line}</text:span>" in out
