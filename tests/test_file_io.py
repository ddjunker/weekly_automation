"""Backups taken before the master is rewritten in place."""

from __future__ import annotations

import os

from scripts.utils.file_io import backup_file


def test_backup_copies_content_into_a_hidden_folder(tmp_path):
    master = tmp_path / "M10-04.md"
    master.write_text("original", encoding="utf-8")

    backup = backup_file(master, tag="text_gather")

    assert backup is not None
    assert backup.parent.name == ".backups", "dot-prefixed so Obsidian ignores it"
    assert backup.read_text(encoding="utf-8") == "original"
    assert master.read_text(encoding="utf-8") == "original"


def test_missing_file_is_not_an_error(tmp_path):
    assert backup_file(tmp_path / "nope.md", tag="text_gather") is None


def test_keeps_only_the_newest_copies(tmp_path):
    master = tmp_path / "M10-04.md"
    for i in range(8):
        master.write_text(f"v{i}", encoding="utf-8")
        backup = backup_file(master, tag="text_gather", keep=5)
        assert backup is not None and backup.exists(), "the new backup must survive pruning"
        os.utime(backup, (1000 + i, 1000 + i))

    kept = sorted((tmp_path / ".backups").glob("*"), key=lambda p: p.stat().st_mtime)
    assert [p.read_text(encoding="utf-8") for p in kept] == ["v3", "v4", "v5", "v6", "v7"]


def test_same_second_runs_do_not_overwrite_each_other(tmp_path):
    master = tmp_path / "M10-04.md"
    master.write_text("first", encoding="utf-8")
    a = backup_file(master, tag="t", keep=0)
    master.write_text("second", encoding="utf-8")
    b = backup_file(master, tag="t", keep=0)

    assert a != b
    assert a.read_text(encoding="utf-8") == "first"
    assert b.read_text(encoding="utf-8") == "second"


def test_tags_are_independent(tmp_path):
    master = tmp_path / "M10-04.md"
    master.write_text("x", encoding="utf-8")
    backup_file(master, tag="text_gather", keep=1)
    backup_file(master, tag="music_gather", keep=1)
    assert len(list((tmp_path / ".backups").glob("*"))) == 2
