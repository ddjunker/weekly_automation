"""The scripture reference parser — one implementation, shared by publish and the gathers."""

from __future__ import annotations

import pytest

from scripts.utils.openlp import (
    fetch_scripture_verses,
    get_scripture_text,
    sanitize_scripture_reference_for_openlp as sanitize,
)


def refs(passage: str) -> list[tuple[int, int]]:
    return [(v.chapter, v.verse) for v in fetch_scripture_verses("elkton", passage)]


def test_single_verse(bible_db):
    assert refs("Genesis 1:1") == [(1, 1)]


def test_simple_range(bible_db):
    assert refs("Genesis 1:1-4") == [(1, 1), (1, 2), (1, 3), (1, 4)]


def test_comma_list_of_ranges(bible_db):
    """This week's first reading shape: Exodus 20:1-4, 7-9, 12-20."""
    assert refs("Genesis 1:1-2, 5, 8-9") == [(1, 1), (1, 2), (1, 5), (1, 8), (1, 9)]


def test_chapter_carries_across_a_comma_list(bible_db):
    assert refs("Genesis 2:9-10, 3:1-2") == [(2, 9), (2, 10), (3, 1), (3, 2)]


def test_range_crossing_a_chapter(bible_db):
    assert refs("Genesis 1:9-2:2") == [(1, 9), (1, 10), (2, 1), (2, 2)]


def test_semicolon_segments(bible_db):
    assert refs("Genesis 1:1; 3:4") == [(1, 1), (3, 4)]


def test_book_name_with_a_numeral(bible_db):
    assert refs("1 John 2:1-2") == [(2, 1), (2, 2)]


def test_verse_part_letters_are_stripped(bible_db):
    """Philippians 3:4b-14 must find verse 4, not fail on the 'b'."""
    assert refs("Genesis 1:4b-6") == [(1, 4), (1, 5), (1, 6)]


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("Genesis 12:1-4a", "Genesis 12:1-4"),
        ("John 3:16b,17", "John 3:16,17"),
        ("Genesis 1:1–3", "Genesis 1:1-3"),   # en dash
        ("Genesis 1:1—3", "Genesis 1:1-3"),   # em dash
    ],
)
def test_sanitize_reference(raw, expected):
    assert sanitize(raw) == expected


def test_unparsable_reference_raises(bible_db):
    with pytest.raises(ValueError, match="Unable to parse"):
        fetch_scripture_verses("elkton", "!!!")


def test_unknown_book_raises(bible_db):
    with pytest.raises(ValueError, match="not found"):
        fetch_scripture_verses("elkton", "Hezekiah 3:1")


def test_bad_verse_number_names_the_reference(bible_db):
    with pytest.raises(ValueError, match="verse number"):
        fetch_scripture_verses("elkton", "Genesis 1:x-3")


def test_text_helper_joins_verses(bible_db):
    assert get_scripture_text("elkton", "Genesis 1:1-2") == "Genesis 1:1\nGenesis 1:2"
