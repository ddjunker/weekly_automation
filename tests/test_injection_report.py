"""The run summary: a step that matched nothing must say so."""

from __future__ import annotations

import scripts.publish as pub
from scripts.utils.report import InjectionReport, render_summary


def inject(osz, church, master, **kwargs) -> InjectionReport:
    return pub._inject_custom_slides_into_openlp_service(osz, church, master, **kwargs)


def findings(report: InjectionReport, slot: str) -> list:
    return [f for f in report.findings if f.slot == slot]


def test_renamed_reading_item_is_an_error(osz_factory, custom_db, master_md):
    """Previously the matcher's return value was discarded and the miss was silent."""

    def rename(data):
        for entry in data:
            svc = entry.get("serviceitem") if isinstance(entry, dict) else None
            if isinstance(svc, dict) and svc["header"].get("title") == "Reading: Gospel":
                svc["header"]["title"] = "Reading: The Gospel"

    master = master_md + "\n{scripture_ref_3}\nMatthew 21:33-46\n"
    report = inject(osz_factory(mutate=rename), "elkton", master)

    gospel = findings(report, "reading_gospel")
    assert gospel and gospel[0].status == "error"
    assert "Reading: Gospel" in gospel[0].detail


def test_single_paragraph_benediction_is_an_error(osz_factory, custom_db, master_md):
    master = master_md.replace(
        "Go now in the righteousness of faith\nand live by God's just demands.\n\n"
        "And may God's perfect word revive your soul;\nMay Christ Jesus be your savior.\n",
        "Only one paragraph here.\n",
    )
    report = inject(osz_factory(), "elkton", master)
    ben = findings(report, "benediction")
    assert ben and ben[0].status == "error"
    assert "two blank-line separated paragraphs" in ben[0].detail


def test_absent_block_is_a_skip_not_an_error(osz_factory, custom_db):
    report = inject(osz_factory(), "elkton", "{communion}\n0\n")
    statuses = {f.slot: f.status for f in report.findings}
    assert statuses["benediction"] == "skip"
    assert statuses["offertory_prayer"] == "skip"
    assert statuses["ctw"] == "skip"


def test_aof_absent_from_a_communion_template_is_a_skip(osz_factory, custom_db, master_md):
    """
    aof_holder only exists in the non-communion Elkton template, so its absence on a
    communion Sunday is normal — flagging it would make the summary noise.
    """
    master = master_md.replace("{aof_ref}\nomit\n", "{aof_ref}\n12top\n")
    report = inject(
        osz_factory("svc_elk_communion1.osj"), "elkton", master,
        template_name="template_communion1.osz",
    )
    aof = findings(report, "aof")
    assert aof and aof[0].status == "skip"
    assert "template_communion1.osz" in aof[0].detail


def test_nothing_matched_marks_the_output_stale(osz_factory, custom_db):
    report = inject(osz_factory(), "elkton", "")
    assert report.changed is False
    assert report.stale is True


def test_summary_names_every_failure():
    report = InjectionReport(church="lb", template="template_lb.osz", communion="0")
    report.ok("ctw", "CtW Psalm 19")
    report.skip("aof", "not in this template")
    report.error("reading_gospel", "no custom item titled 'Reading: Gospel'")
    report.stale = True

    text = render_summary({"lb": report}, {"lb": "Service--2026-10-04.osz"})
    assert "FAILED" in text
    assert "CONTAINS TEMPLATE CONTENT" in text
    assert "reading_gospel" in text
    assert "1 ok / 1 skip / 0 warn / 1 error" in text
    assert "ctw" not in text.split("\n")[-1], "plain successes stay out of the detail lines"


def test_dry_run_writes_nothing(osz_factory, custom_db, master_md, monkeypatch):
    """The .osz on disk must be byte-identical after a dry run."""
    path = osz_factory()
    before = path.read_bytes()
    report = inject(path, "elkton", "{communion}\n0\n")
    assert report.changed is False
    assert path.read_bytes() == before
