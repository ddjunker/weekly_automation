import os
import shutil
import tempfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

"""
file_io.py - UTF-8 file reads and writes.

Plain helpers plus the safe variants: atomic_write / write_text_atomic /
replace_file_atomic put the new bytes in a sibling temp file and os.replace them
into position, and backup_file keeps the last few versions of a file that is
rewritten in place. No text cleaning (text_clean.py) and no path resolution
(config.py) happens here.
"""

def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")

def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


@contextmanager
def atomic_write(path: Path, *, mode: str = "w", encoding: str | None = "utf-8",
                 backup: bool = False, tag: str = "backup"):
    """
    Write through a temp file in the same directory, then os.replace it into place.

    Same directory matters: replace is only atomic within one filesystem, and these
    paths live on a synced folder. If the body raises, the temp file is removed and
    the original is left exactly as it was — the point being that an interrupted run
    can no longer leave a half-written bulletin or master behind.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if backup:
        backup_file(path, tag=tag)

    fd, tmp_name = tempfile.mkstemp(dir=path.parent, prefix=f".{path.name}.", suffix=".tmp")
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        if "b" in mode:
            handle = tmp_path.open(mode)
        else:
            handle = tmp_path.open(mode, encoding=encoding, newline="")
        try:
            yield handle
            handle.flush()
            os.fsync(handle.fileno())
        finally:
            handle.close()
        os.replace(tmp_path, path)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise


def write_text_atomic(path: Path, text: str, *, backup: bool = False, tag: str = "backup") -> None:
    """Replace a text file in one step, optionally backing up what was there."""
    with atomic_write(path, backup=backup, tag=tag) as fh:
        fh.write(text)


def replace_file_atomic(src: Path, dst: Path, *, backup: bool = False, tag: str = "backup",
                        chunk: int = 1 << 20) -> None:
    """
    Copy src over dst through a temp file next to dst, in chunks.

    The chunking is what keeps large .osz files readable across the OneDrive FUSE
    mount's 10MB sequential-read boundary; the temp-then-replace is what stops a
    failure part way through from destroying the previous service file.
    """
    src, dst = Path(src), Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    if backup:
        backup_file(dst, tag=tag)

    fd, tmp_name = tempfile.mkstemp(dir=dst.parent, prefix=f".{dst.name}.", suffix=".tmp")
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        with src.open("rb") as fin, tmp_path.open("wb") as fout:
            while True:
                block = fin.read(chunk)
                if not block:
                    break
                fout.write(block)
            fout.flush()
            os.fsync(fout.fileno())
        os.replace(tmp_path, dst)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise

def backup_file(path: Path, *, tag: str, keep: int = 5) -> Path | None:
    """
    Copy `path` to a timestamped sibling in a hidden `.backups` folder before it is
    rewritten in place, keeping the newest `keep` copies for that file and tag.

    The folder is dot-prefixed so Obsidian does not index the backups as notes.
    Returns the backup path, or None if there was nothing to back up.
    """
    if not path.exists():
        return None

    backup_dir = path.parent / ".backups"
    backup_dir.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = backup_dir / f"{path.stem}.{tag}-{stamp}{path.suffix}"
    # Two runs inside the same second must not overwrite each other's backup.
    for n in range(1, 100):
        if not backup.exists():
            break
        backup = backup_dir / f"{path.stem}.{tag}-{stamp}-{n}{path.suffix}"
    shutil.copy2(path, backup)

    # Prune by mtime, not by name: a same-second counter suffix sorts before the
    # plain timestamp, so name order is not age order. copy2 preserves the source
    # mtime, so ties are possible — never prune the copy just made.
    existing = sorted(
        (p for p in backup_dir.glob(f"{path.stem}.{tag}-*{path.suffix}") if p != backup),
        key=lambda p: p.stat().st_mtime,
    )
    excess = len(existing) - (keep - 1) if keep > 0 else 0
    for stale in existing[:excess] if excess > 0 else []:
        stale.unlink(missing_ok=True)

    return backup
