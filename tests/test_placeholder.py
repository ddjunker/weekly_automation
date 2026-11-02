"""Master-block extraction: the rules that decide where a placeholder's value ends."""

from __future__ import annotations

import pytest

from scripts.utils.placeholder import (
    PlaceholderNotFound,
    append_below_placeholder,
    extract_block,
    has_placeholder,
    write_block,
)


def test_openlp_tag_does_not_end_a_block(master_md):
    """
    Regression: {h}Red Hymnal #395{/h} at the start of a line used to terminate the
    block, cutting a hymn down to its first heading in the pastor's script.
    """
    value = extract_block(master_md, "song_middle_text_lb")
    assert value.startswith("### verse 1")
    assert "{h}Red Hymnal #395{/h}" in value
    assert "Take time to be holy," in value
    assert "### verse 2" in value
    assert "The world rushes on;" in value


@pytest.mark.parametrize(
    "token_line",
    ["{empty_one}", "=={empty_one}==", "{empty_one} %% a note %%", "{empty_one}  "],
)
def test_empty_block_does_not_swallow_the_next_placeholder(token_line):
    """The property the line-anchored lookahead exists for, across the token spellings."""
    md = f"{token_line}\n\n{{next_one}}\nreal value\n"
    assert extract_block(md, "empty_one") == ""
    assert extract_block(md, "next_one") == "real value"


def test_crlf_master_is_handled():
    md = "{a_block}\r\nfirst\r\nsecond\r\n\r\n{b_block}\r\nother\r\n"
    assert extract_block(md, "a_block").splitlines() == ["first", "second"]
    assert extract_block(md, "b_block") == "other"


def test_two_character_tags_do_not_end_a_block():
    """{it} and {st} are OpenLP formatting, not placeholders, even alone on a line."""
    md = "{a_block}\nfirst\n{it}\nstill the same block\n\n{b_block}\nother\n"
    value = extract_block(md, "a_block")
    assert "still the same block" in value
    assert extract_block(md, "b_block") == "other"


def test_three_character_names_are_placeholders():
    """{001} is a real placeholder in the funeral templates — it must still terminate."""
    md = "{a_block}\nfirst\n{001}\nsecond\n"
    assert extract_block(md, "a_block") == "first"


def test_missing_placeholder_extracts_empty():
    assert extract_block("{a_block}\nvalue\n", "nope") == ""
    assert has_placeholder("{a_block}\nvalue\n", "a_block") is True


def test_append_below_placeholder_is_not_idempotent():
    """
    Documents today's behaviour: a second gather run stacks another copy under the
    placeholder, and the absent-placeholder case passes silently. Both are why
    text_gather/music_gather now back the master up before rewriting it.
    """
    md = "{a_block}\n\n{b_block}\nxx\n"
    once = append_below_placeholder(md, "a_block", "VERSE")
    twice = append_below_placeholder(once, "a_block", "VERSE")
    assert extract_block(once, "a_block") == "VERSE"
    assert extract_block(twice, "a_block") == "VERSE\nVERSE"
    assert append_below_placeholder(md, "not_here", "X") == md


def test_write_block_is_idempotent():
    """A second gather run must leave one copy of the text, not two."""
    md = "{a_block}\n\n{b_block}\nkeep me\n"
    once = write_block(md, "a_block", "VERSE ONE")
    twice = write_block(once, "a_block", "VERSE ONE")
    assert extract_block(once, "a_block") == "VERSE ONE"
    assert extract_block(twice, "a_block") == "VERSE ONE"
    assert extract_block(twice, "b_block") == "keep me"


def test_write_block_replaces_existing_content():
    md = "{a_block}\nOLD TEXT\n\n{b_block}\nkeep me\n"
    out = write_block(md, "a_block", "NEW TEXT")
    assert extract_block(out, "a_block") == "NEW TEXT"
    assert extract_block(out, "b_block") == "keep me"


def test_write_block_at_end_of_file():
    md = "{a_block}\nOLD\n"
    assert extract_block(write_block(md, "a_block", "NEW"), "a_block") == "NEW"


def test_write_block_keeps_multiline_content():
    md = "{a_block}\n\n{b_block}\nkeep\n"
    out = write_block(md, "a_block", "### verse 1\n{h}Red Hymnal #395{/h}\nline two")
    assert extract_block(out, "a_block") == "### verse 1\n{h}Red Hymnal #395{/h}\nline two"
    assert extract_block(out, "b_block") == "keep"


def test_write_block_missing_placeholder_raises():
    with pytest.raises(PlaceholderNotFound):
        write_block("{a_block}\nx\n", "not_here", "value")


def test_append_mode_still_accumulates():
    """Diagnostics genuinely want to stack several lines under one placeholder."""
    md = "{a_block}\n\n{b_block}\nx\n"
    once = write_block(md, "a_block", "first", mode="append")
    twice = write_block(once, "a_block", "second", mode="append")
    assert extract_block(twice, "a_block").splitlines() == ["second", "first"]


def test_append_mode_keeps_token_alone_on_its_line():
    """Appended text glued onto the token line let the block above run through it."""
    md = "{ref_block}\nCtW Psalm 99\n{xml_block} %% note %%\n\n{next_block}\nx\n"
    out = write_block(md, "xml_block", "Leader: text", mode="append")
    assert "{xml_block} %% note %%\nLeader: text\n" in out
    assert extract_block(out, "ref_block") == "CtW Psalm 99"
    assert extract_block(out, "xml_block") == "Leader: text"
    assert extract_block(write_block("{a_block}", "a_block", "v", mode="append"), "a_block") == "v"
