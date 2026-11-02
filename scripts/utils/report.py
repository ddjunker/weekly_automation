"""
What a publish run actually did to a service file.

Every injection step records a finding, so a step that matched nothing can no
longer pass unnoticed and leave the template's own content on the projector. The
run still writes its output files — a file you can fix by hand in OpenLP beats no
file — but it ends with an explicit summary of what was missed.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

Status = Literal["ok", "skip", "warn", "error"]

_ORDER: dict[str, int] = {"error": 0, "warn": 1, "skip": 2, "ok": 3}


@dataclass(frozen=True)
class Finding:
    slot: str
    status: Status
    detail: str = ""


@dataclass
class InjectionReport:
    """Findings for one church's service file."""

    church: str
    template: str = ""
    communion: str = ""
    findings: list[Finding] = field(default_factory=list)
    stale: bool = False

    def record(self, slot: str, status: Status, detail: str = "") -> None:
        self.findings.append(Finding(slot, status, detail))

    def ok(self, slot: str, detail: str = "") -> None:
        self.record(slot, "ok", detail)

    def skip(self, slot: str, detail: str = "") -> None:
        self.record(slot, "skip", detail)

    def warn(self, slot: str, detail: str) -> None:
        self.record(slot, "warn", detail)

    def error(self, slot: str, detail: str) -> None:
        self.record(slot, "error", detail)

    @property
    def changed(self) -> bool:
        return any(f.status == "ok" for f in self.findings)

    @property
    def errors(self) -> list[Finding]:
        return [f for f in self.findings if f.status == "error"]

    @property
    def warnings(self) -> list[Finding]:
        return [f for f in self.findings if f.status == "warn"]

    def counts(self) -> dict[str, int]:
        counts = {"ok": 0, "skip": 0, "warn": 0, "error": 0}
        for finding in self.findings:
            counts[finding.status] += 1
        return counts

    def headline(self, destination: str = "") -> str:
        counts = self.counts()
        verdict = "FAILED" if self.errors else ("OK" if self.changed else "NOTHING MATCHED")
        line = (
            f"{self.church:<8} {verdict:<15} "
            f"{counts['ok']} ok / {counts['skip']} skip / "
            f"{counts['warn']} warn / {counts['error']} error"
        )
        if destination:
            line += f"   -> {destination}"
        if self.stale:
            line += "   (CONTAINS TEMPLATE CONTENT)"
        return line

    def render(self, destination: str = "") -> str:
        """Headline plus every finding that is not a plain success."""
        lines = [self.headline(destination)]
        notable = [f for f in self.findings if f.status != "ok"]
        notable.sort(key=lambda f: _ORDER[f.status])
        for finding in notable:
            lines.append(f"   {finding.status.upper():<6} {finding.slot:<20} {finding.detail}")
        return "\n".join(lines)


def render_summary(reports: dict[str, InjectionReport], destinations: dict[str, str] | None = None) -> str:
    """The end-of-run block printed for all churches."""
    destinations = destinations or {}
    if not reports:
        return ""
    out = ["", "=== OpenLP summary ==="]
    for church, report in reports.items():
        out.append(report.render(destinations.get(church, "")))
    return "\n".join(out)
