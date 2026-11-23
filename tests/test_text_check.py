"""prepub's text check: which custom slides it looks for."""

from __future__ import annotations

import pytest

import scripts.text_gather as tg


@pytest.fixture
def searched(monkeypatch):
    """Record every (church, prefix) lookup and report each as one match."""
    calls: list[tuple[str, str]] = []

    def fake_find(church, prefix, *args, **kwargs):
        calls.append((church, prefix))
        return [{"raw_title": prefix, "title": prefix}]

    monkeypatch.setattr(tg, "find_custom_slides", fake_find)
    return calls


def _master(aof: str) -> str:
    return f"{{ctw_ref}}\nPsalm 122:1-9\n\n{{aof_ref}}\n{aof}\n\n{{aof_title}}\n"


def test_aof_checked_for_elkton_only(searched):
    results = tg._check_custom_slides(_master("38bottom"))
    aof = [r for r in results if r["section"] == "AoF"]
    assert [r["church"] for r in aof] == ["elkton"]
    assert ("lb", "AoF p38bottom") not in searched


def test_ctw_still_checked_for_both_churches(searched):
    results = tg._check_custom_slides(_master("38bottom"))
    assert sorted(r["church"] for r in results if r["section"] == "CtW") == ["elkton", "lb"]


@pytest.mark.parametrize("value", ["omit", "OMIT", "delete", "omit\nnotes"])
def test_no_aof_this_week_is_not_searched(searched, value):
    results = tg._check_custom_slides(_master(value))
    assert not [r for r in results if r["section"] == "AoF"]
    assert not [p for _, p in searched if p.startswith("AoF")]
