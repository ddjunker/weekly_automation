"""
Offertory and benediction fetchers, run against saved copies of each site.

The fixtures in tests/data/web are trimmed real responses for November 8, 2026,
so a site redesign shows up here as a failing test once the fixtures are refreshed.
Nothing here touches the network.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from scripts.utils import web_sources as ws

WEB = Path(__file__).parent / "data" / "web"
NOV_8 = date(2026, 11, 8)


def fixture(name: str) -> str:
    return (WEB / name).read_text(encoding="utf-8")


@pytest.fixture
def serve(monkeypatch):
    """Answer _get from fixtures: serve({"url fragment": "fixture file"})."""

    def install(routes: dict[str, str]):
        def fake_get(url: str) -> str:
            for fragment, name in routes.items():
                if fragment in url:
                    return fixture(name)
            raise ws.SourceUnavailable(f"no fixture for {url}")

        monkeypatch.setattr(ws, "_get", fake_get)

    return install


@pytest.mark.parametrize("raw", ["November 8, 2026", "November 8 2026", "Nov 8, 2026", "2026-11-08", " November  8, 2026 "])
def test_parse_cal_date(raw):
    assert ws.parse_cal_date(raw) == NOV_8


def test_parse_cal_date_unreadable():
    assert ws.parse_cal_date("the second Sunday") is None


@pytest.mark.parametrize("url, fetcher", [
    ("https://umfmichigan.org/weekly-offertory-prayer/", ws.fetch_umf_offertory),
    ("https://laughingbird.net/occasions/", ws.fetch_laughingbird_benediction),
    ("https://www.umcdiscipleship.org/equipping-leaders/stewardship", ws.fetch_umc_offertory),
    ("https://example.org/prayers", None),
])
def test_fetcher_for(url, fetcher):
    assert ws.fetcher_for(url) is fetcher


# ---------------- LaughingBird ----------------

def test_laughingbird_keeps_leader_lines_only(serve):
    serve({"laughingbird.net": "laughingbird_2026-11-08.json"})
    assert ws.fetch_laughingbird_benediction(NOV_8) == (
        "Go now, and follow in the footsteps\n"
        "of those who have gone before us in Christ.\n"
        "Revere the Lord.\n"
        "Serve God in sincerity and faithfulness.\n"
        "Teach those who come after you to trust in God.\n"
        "\n"
        "And may God protect you along the way;\n"
        "May Christ Jesus keep you alert and prepared;\n"
        "And may the Holy Spirit fuel the lamp that guides your path."
    )


def test_laughingbird_wrong_date(serve):
    serve({"laughingbird.net": "laughingbird_2026-11-08.json"})
    with pytest.raises(ws.SourceUnavailable):
        ws.fetch_laughingbird_benediction(date(2026, 11, 15))


def test_laughingbird_missing_section():
    with pytest.raises(ws.SourceUnavailable):
        ws.extract_laughingbird_benediction("<p>Prayers of the People</p>")


# ---------------- UMF ----------------

def test_umf_picks_the_post_titled_by_date(serve):
    """The search also returns a newsletter; only the dated post counts."""
    serve({"umfmichigan.org": "umf_posts_2026-11-08.json"})
    text = ws.fetch_umf_offertory(NOV_8)
    assert text.startswith("God of Grace, your promise of a heavenly kingdom inspires us.")
    assert text.endswith("We pray in Jesus' name. Amen.")
    assert "Matthew" not in text


def test_umf_no_post(serve):
    serve({"umfmichigan.org": "umf_posts_2026-11-08.json"})
    with pytest.raises(ws.SourceUnavailable):
        ws.fetch_umf_offertory(date(2026, 11, 15))


# ---------------- UMC Discipleship ----------------

UMC_ROUTES = {
    "-offertory-prayer": "umc_offertory_2026-11-08.html",
    "/worship-planning/": "umc_week_2026-11-08.html",
    "umcdiscipleship.org/": "umc_home.html",
}


def test_umc_follows_homepage_to_offertory(serve):
    serve(UMC_ROUTES)
    text = ws.fetch_umc_offertory(NOV_8)
    assert text.startswith("Watchful God, you call us to keep awake")
    assert text.endswith("faithful in your service, through Christ our Lord. Amen.")


def test_umc_homepage_week_url():
    assert ws.extract_umc_week_url(fixture("umc_home.html"), NOV_8).endswith(
        "/it-will-be-like-this/week-1-twenty-fourth-sunday-after-pentecost-year-a"
    )


def test_umc_sunday_not_listed():
    with pytest.raises(ws.SourceUnavailable):
        ws.extract_umc_week_url(fixture("umc_home.html"), date(2026, 12, 6))


def test_umc_offertory_page_for_another_week_is_refused():
    with pytest.raises(ws.SourceUnavailable, match="dated November 8"):
        ws.extract_umc_offertory(fixture("umc_offertory_2026-11-08.html"), date(2026, 11, 15))


# The homepage moves the coming Sunday into a featured block
# (<div class="date-box fixed-width"> + <h2><a href>) and drops it from
# "Upcoming Events"; its week and offertory pages badge it the same way.
NOV_22 = date(2026, 11, 22)
UMC_ROUTES_FEATURED = {
    "-offertory-prayer": "umc_offertory_2026-11-22.html",
    "/worship-planning/": "umc_week_2026-11-22.html",
    "umcdiscipleship.org/": "umc_home_2026-11-16.html",
}


def test_umc_featured_sunday_followed_to_offertory(serve):
    serve(UMC_ROUTES_FEATURED)
    text = ws.fetch_umc_offertory(NOV_22)
    assert text.startswith("Compassionate Shepherd, you have blessed us")
    assert text.endswith("to love as Christ first loved us. Amen.")


@pytest.mark.parametrize("day, slug", [
    (NOV_22, "/week-3-twenty-sixth-sunday-after-pentecost-year-a"),
    (date(2026, 11, 15), "/week-2-twenty-fifth-sunday-after-pentecost-year-a"),
    (NOV_8, "/week-1-twenty-fourth-sunday-after-pentecost-year-a"),
])
def test_umc_each_badge_gets_its_own_link(day, slug):
    url = ws.extract_umc_week_url(fixture("umc_home_2026-11-16.html"), day)
    assert url.endswith(slug)
