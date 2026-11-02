"""
Regenerate the tests/data/*.osj fixtures from the real OpenLP templates.

Not part of the test run — run it by hand when a template changes:

    python3 tests/tools/make_osz_fixture.py

service_data.osj is ~35 KB inside a 120-225 MB template, and zipfile reads it via
the central directory without touching the embedded videos, so this is quick. Slide
bodies are blanked: the tests care about item structure (plugin, title, footer),
not about shipping copyrighted lyrics and liturgy into the repo.
"""

from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.utils.openlp import get_openlp_root  # noqa: E402

DATA = Path(__file__).resolve().parents[1] / "data"

SOURCES = {
    "svc_elk_sunday.osj": ("elkton", "template_sunday.osz"),
    "svc_elk_communion1.osj": ("elkton", "template_communion1.osz"),
    "svc_lb.osj": ("lb", "template_lb.osz"),
}


def trim(service_data: list) -> list:
    """Keep headers intact; replace slide bodies with a short marker."""
    for entry in service_data:
        svc = entry.get("serviceitem") if isinstance(entry, dict) else None
        if not isinstance(svc, dict):
            continue
        slides = svc.get("data")
        if isinstance(slides, list):
            for idx, slide in enumerate(slides, start=1):
                if isinstance(slide, dict) and "raw_slide" in slide:
                    slide["raw_slide"] = f"[fixture slide {idx}]"
    return service_data


def main() -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    for out_name, (church, template) in SOURCES.items():
        src = get_openlp_root(church) / "OpenLP Services" / template
        if not src.exists():
            print(f"[SKIP] {template} not found at {src}")
            continue
        with zipfile.ZipFile(src) as zf:
            data = json.loads(zf.read("service_data.osj").decode("utf-8"))
        (DATA / out_name).write_text(
            json.dumps(trim(data), ensure_ascii=False, indent=1), encoding="utf-8"
        )
        print(f"[OK] {out_name} ← {template} ({len(data)} items)")


if __name__ == "__main__":
    main()
