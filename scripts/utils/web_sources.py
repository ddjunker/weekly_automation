"""
Fetch offertory and benediction text straight from the source sites.

text_gather used to open each source in a browser and wait for a manual copy.
For the sites below the text can be read directly, so the browser is only
needed when a fetch fails (site down, layout changed, week not posted yet).

Each fetcher takes the service date and returns cleaned text, or raises
SourceUnavailable with a reason the caller can show.
"""

from __future__ import annotations

import html
import json
import re
from datetime import date, datetime
from urllib.error import URLError
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

from .text_clean import clean_text

USER_AGENT = "Mozilla/5.0 (weekly_automation)"
TIMEOUT = 20


class SourceUnavailable(Exception):
    """The source could not supply text for this date."""


# ---------------- Helpers ----------------

def parse_cal_date(raw: str) -> date | None:
    """Parse a master's {cal_date} value ("November 8, 2026" and friends)."""
    raw = clean_text(raw or "")
    for fmt in ("%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def _get(url: str) -> str:
    req = Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urlopen(req, timeout=TIMEOUT) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except (URLError, TimeoutError, OSError) as e:
        raise SourceUnavailable(f"could not fetch {url}: {e}") from e


def _get_json(url: str):
    try:
        return json.loads(_get(url))
    except json.JSONDecodeError as e:
        raise SourceUnavailable(f"unexpected response from {url}") from e


def _paragraphs(fragment: str) -> list[str]:
    """Raw inner HTML of each <p> in a fragment."""
    return re.findall(r"(?s)<p\b[^>]*>(.*?)</p>", fragment)


def _html_to_text(fragment: str) -> str:
    """One HTML paragraph to text: <br> becomes a line break, tags drop out."""
    text = re.sub(r"(?i)<br\s*/?>", "\n", fragment)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text).replace("\xa0", " ")
    lines = [line.strip() for line in text.split("\n")]
    return clean_text("\n".join(line for line in lines if line))


def _join(paragraphs: list[str]) -> str:
    return "\n\n".join(p for p in (_html_to_text(p) for p in paragraphs) if p)


def _long_date(day: date) -> str:
    return f"{day:%B} {day.day}, {day.year}"


# ---------------- United Methodist Foundation of Michigan ----------------

UMF_POSTS_API = "https://umfmichigan.org/wp-json/wp/v2/posts"


def fetch_umf_offertory(day: date) -> str:
    """Weekly offertory prayer. Posts are titled by date: "November 8,  2026"."""
    title = _long_date(day)
    query = urlencode({"search": title, "_fields": "title,content", "per_page": 20})
    posts = _get_json(f"{UMF_POSTS_API}?{query}")

    def norm(s: str) -> str:
        return re.sub(r"\s+", " ", html.unescape(s)).strip()

    for post in posts if isinstance(posts, list) else []:
        if norm(post.get("title", {}).get("rendered", "")) != title:
            continue
        # The prayer is the first paragraph; the lectionary reference follows it.
        for para in _paragraphs(post.get("content", {}).get("rendered", "")):
            text = _html_to_text(para)
            if text:
                return text
        raise SourceUnavailable(f"UMF post for {title} has no prayer text")
    raise SourceUnavailable(f"no UMF post titled {title}")


# ---------------- LaughingBird ----------------

LAUGHINGBIRD_EVENTS_API = "https://laughingbird.net/wp-json/tribe/events/v1/events"


def extract_laughingbird_benediction(description: str) -> str:
    """
    The "Commission & Benediction" accordion section from an occasion page.

    Keeps the leader's paragraphs. Drops the people's response (the paragraph
    carrying the italic "In the name of Christ. Amen.") and the copyright line.
    """
    m = re.search(
        r'(?s)data-title="efaq-commission[^"]*".*?<div class="ea-content-wrapper[^"]*">(.*?)</div>',
        description,
    )
    if not m:
        raise SourceUnavailable("no Commission & Benediction section")
    keep = [
        p for p in _paragraphs(m.group(1))
        if "©" not in p and not re.search(r"(?i)<(i|em)\b", p)
    ]
    text = _join(keep)
    if not text:
        raise SourceUnavailable("Commission & Benediction section is empty")
    return text


def fetch_laughingbird_benediction(day: date) -> str:
    iso = day.isoformat()
    query = urlencode({"start_date": iso, "end_date": iso, "per_page": 5})
    data = _get_json(f"{LAUGHINGBIRD_EVENTS_API}?{query}")
    events = data.get("events", []) if isinstance(data, dict) else []
    for event in events:
        if str(event.get("start_date", "")).startswith(iso):
            return extract_laughingbird_benediction(event.get("description", ""))
    raise SourceUnavailable(f"no LaughingBird occasion on {iso}")


# ---------------- UMC Discipleship Ministries ----------------

UMC_HOME = "https://www.umcdiscipleship.org/"

# The date badge on the homepage and on each worship page:
#   <div class="date-box"><span>8</span><p class="full">November 2026</p>...
# The featured (current) week and a page's own badge carry an extra class:
#   <div class="date-box fixed-width">...
_UMC_DATE_BOX = (r'<div class="date-box(?:\s[^"]*)?">\s*<span>(\d+)</span>'
                 r'\s*<p class="full">([^<]+)</p>')


def _umc_date_matches(day_num: str, month_year: str, day: date) -> bool:
    return int(day_num) == day.day and month_year.strip() == f"{day:%B} {day.year}"


def extract_umc_week_url(home_html: str, day: date) -> str:
    """The worship-planning page the homepage lists for this Sunday.

    The Sunday may sit in the featured block (title link in an <h2>) or in
    "Upcoming Events" (link with class calendar_event_title). Either way the
    link is the first worship-planning href after its date badge, looked for
    only up to the next badge so one Sunday never borrows another's link.
    """
    badges = list(re.finditer(_UMC_DATE_BOX, home_html))
    for i, m in enumerate(badges):
        if not _umc_date_matches(m.group(1), m.group(2), day):
            continue
        end = badges[i + 1].start() if i + 1 < len(badges) else len(home_html)
        link = re.search(r'<a href="([^"]*/worship-planning/[^"]+)"',
                         home_html[m.end():end])
        if link:
            return html.unescape(link.group(1))
    raise SourceUnavailable(f"UMC homepage lists no worship plan for {_long_date(day)}")


def extract_umc_offertory_url(week_html: str) -> str:
    """The "Offertory Prayer" link in the worship page's sidebar."""
    m = re.search(r'<a[^>]+href="([^"]+)"[^>]*>\s*Offertory Prayer', week_html)
    if not m:
        raise SourceUnavailable("worship page has no Offertory Prayer link")
    return html.unescape(m.group(1))


def extract_umc_offertory(page_html: str, day: date) -> str:
    badge = re.search(_UMC_DATE_BOX, page_html)
    if not badge:
        raise SourceUnavailable("offertory page has no date to check against")
    if not _umc_date_matches(badge.group(1), badge.group(2), day):
        raise SourceUnavailable(
            f"offertory page is dated {badge.group(2).split()[0]} {badge.group(1)}, "
            f"not {_long_date(day)}"
        )
    m = re.search(r"(?s)<h2>\s*Offertory Prayer\s*</h2>(.*?)</div>", page_html)
    if not m:
        raise SourceUnavailable("offertory page has no Offertory Prayer section")
    # Drop the trailing scripture citation: "... Amen. (<a ...>Matthew 25:1-13</a>)"
    paras = [re.sub(r"(?s)\s*\(\s*<a\b[^>]*>.*?</a>\s*\)\s*$", "", p.strip())
             for p in _paragraphs(m.group(1))]
    text = _join(paras)
    if not text:
        raise SourceUnavailable("Offertory Prayer section is empty")
    return text


def fetch_umc_offertory(day: date) -> str:
    week_url = extract_umc_week_url(_get(UMC_HOME), day)
    offertory_url = extract_umc_offertory_url(_get(week_url))
    return extract_umc_offertory(_get(offertory_url), day)


# ---------------- Dispatch ----------------

FETCHERS = {
    "umfmichigan.org": fetch_umf_offertory,
    "laughingbird.net": fetch_laughingbird_benediction,
    "umcdiscipleship.org": fetch_umc_offertory,
}


def fetcher_for(url: str):
    """The fetcher for a source URL's site, or None when it needs manual capture."""
    host = (urlparse(url).hostname or "").lower().removeprefix("www.")
    return FETCHERS.get(host)
