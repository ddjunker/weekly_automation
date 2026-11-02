"""CCLI praise-song lyrics as they land in the master."""

from __future__ import annotations

import pytest

from scripts.music_gather import label_ccli_song_parts


@pytest.mark.parametrize(
    "label",
    ["Verse 1", "Verse", "Chorus", "Chorus 2", "Pre-Chorus", "Pre-Chorus 1", "Bridge",
     "Ending", "Tag", "Interlude 2", "Intro", "BRIDGE", "VERSE 2", "Chorus:"],
)
def test_part_labels_become_headings(label):
    assert label_ccli_song_parts(f"Title\n\n{label}\nA lyric line") == (
        f"Title\n\n### {label}\nA lyric line"
    )


@pytest.mark.parametrize(
    "line", ["Chorus of angels singing", "Verse by verse", "Tagged along", "Ending 1 2"],
)
def test_lyric_lines_starting_with_a_part_name_are_left_alone(line):
    assert label_ccli_song_parts(line) == line


def test_crlf_input_and_indented_label():
    assert label_ccli_song_parts("  Chorus\r\nSing\r\n") == "### Chorus\nSing\n"
