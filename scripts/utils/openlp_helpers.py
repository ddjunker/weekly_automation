import sqlite3
from pathlib import Path
from .text_clean import clean_text

def load_openlp_song(db_path: Path, uuid: int) -> dict | None:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT title, lyrics FROM songs WHERE id = ?", (uuid,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return None
    return {"title": clean_text(row[0] or ""), "lyrics": clean_text(row[1] or "")}

def load_openlp_bible_text(db_path: Path, reference: str) -> str | None:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT scripture FROM bible WHERE ref = ?", (reference,))
    row = cur.fetchone()
    conn.close()
    return clean_text(row[0]) if row else None

def load_openlp_custom_slide(db_path: Path, uuid: int) -> str | None:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT text FROM custom_slide WHERE id=?", (uuid,))
    row = cur.fetchone()
    conn.close()
    return clean_text(row[0]) if row else None

def build_song_index(db_path: Path) -> list[dict]:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
        SELECT songs.id, songs.title, ss.entry, sb.name
        FROM songs
        JOIN songs_songbooks ss ON songs.id = ss.song_id
        JOIN songbooks sb ON ss.songbook_id = sb.id
    """)
    rows = []
    for uuid, title, entry, hymnal in cur.fetchall():
        rows.append({
            "uuid": uuid,
            "title": clean_text(title),
            "entry": str(entry).strip(),
            "hymnal": clean_text(hymnal),
        })
    conn.close()
    return rows

def map_song_index(rows: list[dict]) -> dict:
    imap = {}
    for r in rows:
        key = (clean_text(r["hymnal"]), clean_text(r["entry"]))
        if key in imap:
            continue
        imap[key] = r
    return imap
