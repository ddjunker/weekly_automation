"""
Reading, matching and rewriting items inside an OpenLP service file (.osz).

A service file is a zip whose service_data.osj holds a list of entries. Most are
{"serviceitem": {"header": {...}, "data": [slides]}}; the first is a bare
{"openlp_core": {...}} and malformed entries turn up in hand-edited files, so
every walk has to tolerate both.

This module exists because that walk was written out nine times in publish.py,
each copy slightly different about stripping and case — which is how an item
titled "Offertory Prayer - GBOD" came to be missed by a matcher looking for
"Offertory Prayer".
"""

from __future__ import annotations

import json
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Iterator, Literal, NamedTuple

MatchMode = Literal["exact", "prefix", "any"]

SERVICE_DATA_NAME = "service_data.osj"


class ServiceItem(NamedTuple):
    """One well-formed item, with the dicts the caller needs to mutate in place."""

    index: int
    entry: dict
    svc: dict
    header: dict
    title: str      # original case, whitespace collapsed
    plugin: str     # lowercased


def normalize_title(value: object) -> str:
    """
    Whitespace-collapsed casefold, for comparing service item titles.

    Deliberately does not drop commas: that belongs to searching the custom slide
    library, where titles carry scripture references, not to matching the items
    inside a service file.
    """
    return " ".join(str(value or "").split()).casefold()


def iter_service_items(service_data: list, *, plugin: str | None = None) -> Iterator[ServiceItem]:
    """Yield well-formed service items, optionally only those of one plugin."""
    wanted = plugin.strip().lower() if plugin else None
    for index, entry in enumerate(service_data or []):
        if not isinstance(entry, dict):
            continue
        svc = entry.get("serviceitem")
        if not isinstance(svc, dict):
            continue
        header = svc.get("header")
        if not isinstance(header, dict):
            continue
        item_plugin = str(header.get("plugin", "") or "").strip().lower()
        if wanted is not None and item_plugin != wanted:
            continue
        yield ServiceItem(
            index=index,
            entry=entry,
            svc=svc,
            header=header,
            title=" ".join(str(header.get("title", "") or "").split()),
            plugin=item_plugin,
        )


def _matches(item: ServiceItem, title: str | None, match: MatchMode) -> bool:
    if title is None or match == "any":
        return True
    needle = normalize_title(title)
    hay = normalize_title(item.title)
    return hay.startswith(needle) if match == "prefix" else hay == needle


def find_service_items(
    service_data: list,
    *,
    plugin: str | None = None,
    title: str | None = None,
    match: MatchMode = "exact",
) -> list[ServiceItem]:
    """All items matching the plugin and title predicate, in service order."""
    return [i for i in iter_service_items(service_data, plugin=plugin) if _matches(i, title, match)]


def find_service_item(
    service_data: list,
    *,
    plugin: str | None = None,
    title: str | None = None,
    match: MatchMode = "exact",
) -> ServiceItem | None:
    """The first matching item, or None."""
    for item in iter_service_items(service_data, plugin=plugin):
        if _matches(item, title, match):
            return item
    return None


def has_service_item(
    service_data: list,
    *,
    plugin: str | None = None,
    title: str | None = None,
    match: MatchMode = "exact",
) -> bool:
    return find_service_item(service_data, plugin=plugin, title=title, match=match) is not None


def remove_service_items(
    service_data: list,
    *,
    plugin: str | None = None,
    title: str | None = None,
    match: MatchMode = "exact",
) -> int:
    """Remove matching items in place. Returns how many were removed."""
    doomed = find_service_items(service_data, plugin=plugin, title=title, match=match)
    for item in reversed(doomed):
        del service_data[item.index]
    return len(doomed)


def load_service_data(osz_path: Path) -> list:
    """Read service_data.osj from an OpenLP .osz file and return the parsed list."""
    with zipfile.ZipFile(osz_path, "r") as zin:
        if SERVICE_DATA_NAME not in zin.namelist():
            raise ValueError(f"{SERVICE_DATA_NAME} not found in {osz_path}")
        raw = zin.read(SERVICE_DATA_NAME)
    return json.loads(raw.decode("utf-8"))


def write_service_data(osz_path: Path, service_data: list) -> None:
    """
    Write service_data.osj back into an .osz, preserving every other member.

    osz_path must already be on a local filesystem (e.g. /tmp) — never a FUSE
    mount — so zipfile can read the large AVI entries without hitting the
    OneDrive 10MB sequential-read boundary. The zip is rebuilt in a sibling temp
    file and then replaces osz_path, so an interrupted run cannot leave a
    half-written service file behind.
    """
    new_json = json.dumps(service_data, ensure_ascii=False).encode("utf-8")
    with tempfile.NamedTemporaryFile(suffix=".osz", delete=False, dir=osz_path.parent) as tf:
        tmp_path = Path(tf.name)
    try:
        with zipfile.ZipFile(osz_path, "r") as zin, zipfile.ZipFile(tmp_path, "w") as zout:
            for item in zin.infolist():
                if item.filename == SERVICE_DATA_NAME:
                    zout.writestr(item, new_json)
                else:
                    with zin.open(item.filename) as src, zout.open(item, "w") as dst:
                        shutil.copyfileobj(src, dst)
        tmp_path.replace(osz_path)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
