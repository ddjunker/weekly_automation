"""Song ids and verse selection — one parser for the script and the projector."""

from __future__ import annotations

import pytest

from scripts.utils.hymnal import (
    hymnal_name,
    intro_slide_prefix,
    is_skip_marker,
    parse_song_ref,
    parse_verse_list,
    read_verse_filter,
)


@pytest.mark.parametrize(
    "raw, prefix, entry",
    [
        ("r 123", "r", "123"),          # the documented spaced form publish used to reject
        ("R395", "r", "395"),
        ("b218", "b", "218"),
        ("B218 vv 1-3,5", "b", "218"),
        ("k 42 vv 1, 3, 5-7", "k", "42"),
        ("  r  7  ", "r", "7"),
        ("R395 %% a note %%", "r", "395"),
    ],
)
def test_hymnal_ids_that_must_parse(raw, prefix, entry):
    ref = parse_song_ref(raw)
    assert ref is not None, raw
    assert (ref.prefix, ref.entry) == (prefix, entry)


@pytest.mark.parametrize("raw", ["MercyMe", "", None, "tbd", "omit", "congregational choice", "x12"])
def test_values_that_are_not_hymnal_ids(raw):
    assert parse_song_ref(raw) is None


@pytest.mark.parametrize("raw", ["tbd", "TBD", "omit", "na", "congregational choice", "  tbd  "])
def test_skip_markers(raw):
    assert is_skip_marker(raw) is True


def test_skip_marker_is_not_every_value():
    assert is_skip_marker("R395") is False
    assert is_skip_marker("I Can Only Imagine") is False


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("1, 3", (1, 3)),
        ("1-3,5", (1, 2, 3, 5)),      # ranges expand; the old parser dropped verse 2
        ("1 3 5", (1, 3, 5)),
        ("2-4", (2, 3, 4)),
        ("4-2", (2, 3, 4)),           # tolerate a reversed range
        ("3, 3, 1", (3, 1)),          # de-duplicated, order preserved
        ("omit", None),
        ("all", None),
        ("", None),
        (None, None),
    ],
)
def test_verse_lists(raw, expected):
    assert parse_verse_list(raw) == expected


def test_legacy_verse_suffix_is_parsed_but_separate():
    """The id still identifies the hymn; the verse list is reported, not applied."""
    ref = parse_song_ref("B218 vv 1-3,5")
    assert ref is not None and ref.entry == "218"
    assert ref.legacy_verses == (1, 2, 3, 5)


def test_verse_filter_comes_from_the_master_field():
    md = "{song_middle_vv_elkton}\n1, 3\n\n{song_middle_vv_lb}\nomit\n"
    assert read_verse_filter(md, "middle", "elkton") == (1, 3)
    assert read_verse_filter(md, "middle", "lb") is None
    assert read_verse_filter(md, "opening", "elkton") is None


@pytest.mark.parametrize(
    "church, prefix, expected",
    [
        ("elkton", "r", "United Methodist Hymnal"),
        ("elkton", "b", "The Singing Church"),
        ("lb", "b", "Hymns of Faith"),
        ("lb", "k", "The Faith We Sing"),
        ("elkton", "k", None),      # Elkton has no black hymnal
        ("nowhere", "r", None),
    ],
)
def test_hymnal_names(church, prefix, expected):
    assert hymnal_name(church, prefix) == expected


@pytest.mark.parametrize(
    "church, raw, title, expected",
    [
        ("elkton", "r395", "", "U395"),
        ("elkton", "b18", "", "S018"),
        ("lb", "b18", "", "H018"),
        ("lb", "k7", "", "F0007"),
        ("elkton", "MercyMe", "I Can Only Imagine", "P901 I Can Only Imagine"),
        ("elkton", "MercyMe", "", None),
    ],
)
def test_intro_slide_prefix(church, raw, title, expected):
    assert intro_slide_prefix(church, parse_song_ref(raw), title) == expected
