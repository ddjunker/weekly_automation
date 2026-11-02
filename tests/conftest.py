"""
Shared fixtures.

Nothing here touches the live OneDrive tree or the real OpenLP databases: service
files are rebuilt as a few-KB .osz from the trimmed .osj fixtures in tests/data,
and the custom-slide database is created fresh in tmp_path.
"""

from __future__ import annotations

import json
import sqlite3
import zipfile
from pathlib import Path

import pytest

DATA = Path(__file__).parent / "data"

DEFAULT_OSJ = "svc_elk_sunday.osj"


@pytest.fixture
def service_data():
    """Load a template's service_data.osj as a Python list."""

    def load(osj: str = DEFAULT_OSJ) -> list:
        return json.loads((DATA / osj).read_text(encoding="utf-8"))

    return load


@pytest.fixture
def osz_factory(tmp_path, service_data):
    """
    Build a small .osz from a stored .osj.

    A junk member is always included so tests can prove the rewriter preserves
    everything that is not service_data.osj (the real files carry ~200 MB of video).
    """

    def build(osj: str = DEFAULT_OSJ, *, mutate=None, name: str = "svc.osz") -> Path:
        data = service_data(osj)
        if mutate is not None:
            mutate(data)
        path = tmp_path / name
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("service_data.osj", json.dumps(data, ensure_ascii=False))
            zf.writestr("deadbeef.avi", b"not-a-real-video")
        return path

    return build


@pytest.fixture
def custom_db(tmp_path, monkeypatch):
    """
    A throwaway custom.sqlite with OpenLP's real schema, wired in place of the
    church databases. Returns a helper to seed and to read rows back.
    """
    import scripts.utils.openlp as openlp

    paths: dict[str, Path] = {}

    def db_for(church: str) -> Path:
        path = paths.get(church)
        if path is None:
            path = tmp_path / f"custom-{church}.sqlite"
            conn = sqlite3.connect(path)
            conn.execute(
                "CREATE TABLE custom_slide (id INTEGER NOT NULL, title VARCHAR(255) NOT NULL, "
                "text TEXT NOT NULL, credits TEXT, theme_name VARCHAR(128), PRIMARY KEY (id))"
            )
            conn.commit()
            conn.close()
            paths[church] = path
        return path

    monkeypatch.setattr(openlp, "get_custom_db", db_for)

    class Helper:
        def add(self, church: str, title: str, credits: str = "", text: str = "") -> int:
            conn = sqlite3.connect(db_for(church))
            cur = conn.execute(
                "INSERT INTO custom_slide (title, text, credits, theme_name) VALUES (?, ?, ?, '')",
                (title, text, credits),
            )
            conn.commit()
            row_id = int(cur.lastrowid)
            conn.close()
            return row_id

        def rows(self, church: str) -> list[tuple]:
            conn = sqlite3.connect(db_for(church))
            out = conn.execute(
                "SELECT id, title, IFNULL(credits, ''), text FROM custom_slide ORDER BY id"
            ).fetchall()
            conn.close()
            return out

    return Helper()


@pytest.fixture(autouse=True)
def no_live_databases(monkeypatch):
    """
    Keep the suite off the real OpenLP song and Bible databases: tests must not
    depend on what happens to be installed on one machine. The custom-slide
    database has its own fixture; ask for `custom_db` where a test needs it.
    """
    import scripts.publish as pub
    import scripts.utils.openlp as openlp

    verses = [(1, 1, "In the beginning God created the heavens and the earth.")]
    monkeypatch.setattr(pub, "_fetch_scripture_verses", lambda church, passage: list(verses))
    monkeypatch.setattr(openlp, "get_scripture_text", lambda church, passage: verses[0][2])
    monkeypatch.setattr(openlp, "build_song_index", lambda church: {})


@pytest.fixture
def bible_db(tmp_path, monkeypatch):
    """
    A small Bible database with OpenLP's schema: Genesis 1-3 and 1 John 2, enough
    to exercise ranges, comma lists and chapter transitions without the real file.
    """
    import scripts.utils.openlp as openlp

    path = tmp_path / "bible.sqlite"
    conn = sqlite3.connect(path)
    conn.execute(
        "CREATE TABLE book (id INTEGER PRIMARY KEY, book_reference_id INTEGER, "
        "testament_reference_id INTEGER, name VARCHAR(50))"
    )
    conn.execute(
        "CREATE TABLE verse (id INTEGER PRIMARY KEY, book_id INTEGER, chapter INTEGER, "
        "verse INTEGER, text TEXT)"
    )
    conn.execute("INSERT INTO book (id, name) VALUES (1, 'Genesis'), (2, '1 John')")
    rows = [
        (1, chapter, verse, f"Genesis {chapter}:{verse}")
        for chapter in (1, 2, 3)
        for verse in range(1, 11)
    ] + [(2, 2, verse, f"1 John 2:{verse}") for verse in range(1, 6)]
    conn.executemany(
        "INSERT INTO verse (book_id, chapter, verse, text) VALUES (?, ?, ?, ?)", rows
    )
    conn.commit()
    conn.close()

    monkeypatch.setattr(openlp, "get_bible_db", lambda church: path)
    # fetch_scripture_verses is the real thing here; undo the blanket stub.
    monkeypatch.setattr(
        openlp, "get_scripture_text",
        lambda church, passage: "\n".join(
            v.text for v in openlp.fetch_scripture_verses(church, passage)
        ),
    )
    return path


@pytest.fixture
def master_md() -> str:
    """
    A minimal master, including the two shapes that used to break extraction:
    an OpenLP {h} tag inside a song block, and an 'omit' marker with text under it.
    """
    return (
        "# GenInfo\n\n"
        "{communion}\n0\n\n"
        "{announce_events_elkton}\n"
        "Today after worship - meeting\n"
        "Oct 7th 7:00 PM - UPW\n\n"
        "{announce_events_lb}\nomit  \n\nThanks\n\n"
        "{aof_ref}\nomit\n\n"
        "{ctw_xml_elkton}\n\n"
        "Leader: The heavens are telling the glory of God;\n"
        "and the firmament proclaims his handiwork.\n"
        "People: Day to day pours forth speech, and night to night declares knowledge.\n\n"
        "L: There is no speech;\n"
        "P: yet their voice goes out through all the earth.\n\n"
        "{song_middle_text_lb}\n\n"
        "### verse 1\n"
        "{h}Red Hymnal #395{/h}\n"
        "Take time to be holy,\n"
        "Speak oft with thy Lord;\n\n"
        "### verse 2\n"
        "Take time to be holy,\n"
        "The world rushes on;\n\n"
        "{benediction_text}\n\n"
        "Go now in the righteousness of faith\n"
        "and live by God's just demands.\n\n"
        "And may God's perfect word revive your soul;\n"
        "May Christ Jesus be your savior.\n"
    )
