from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any


class Status(str, Enum):
    CONFORME = "conforme"
    ANOMALIES = "anomalies"
    FICHIER_MANQUANT = "fichier_manquant"
    SANS_REFERENCE = "sans_reference"
    ERREUR = "erreur"


@dataclass(slots=True)
class Anomaly:
    category: str
    code: str
    message: str
    sheet: str | None = None
    element: str | None = None
    expected_position: str | int | None = None
    observed_position: str | int | None = None
    severity: str = "warning"
    impact: int = 1
    details: dict[str, Any] = field(default_factory=dict)
    location: str | None = None
    expected: Any = None
    observed: Any = None
    id: str | None = None
    confidence: Any = None
    consequences: list[dict[str, Any]] = field(default_factory=list)
    action: dict[str, Any] = field(default_factory=dict)
    noise_reduction: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class CountryResult:
    key: str
    country: str
    reference_path: str | None = None
    received_path: str | None = None
    status: Status = Status.CONFORME
    anomalies: list[Anomaly] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    sheet_order_expected: list[str] = field(default_factory=list)
    sheet_order_observed: list[str] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def total_anomalies(self) -> int:
        return sum(max(0, anomaly.impact) for anomaly in self.anomalies)

    @property
    def counts(self) -> dict[str, int]:
        counts = {
            "feuilles": 0,
            "colonnes": 0,
            "lignes": 0,
            "formules": 0,
            "dependances": 0,
            "valeurs": 0,
            "fichier": 0,
            "autres": 0,
        }
        for anomaly in self.anomalies:
            counts[anomaly.category] = counts.get(anomaly.category, 0) + max(0, anomaly.impact)
        return counts

    @property
    def counts_by_code(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for anomaly in self.anomalies:
            counts[anomaly.code] = counts.get(anomaly.code, 0) + max(0, anomaly.impact)
        return counts

    @property
    def root_counts_by_code(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for anomaly in self.anomalies:
            counts[anomaly.code] = counts.get(anomaly.code, 0) + 1
        return counts

    @property
    def impact_summary(self) -> dict[str, int]:
        summary = {"confirmed": 0, "probable": 0, "possible": 0}
        ranks = {"possible": 0, "probable": 1, "confirmed": 2}
        identified_targets: dict[str, str] = {}
        unidentified = {"confirmed": 0, "probable": 0, "possible": 0}
        for anomaly in self.anomalies:
            for consequence in anomaly.consequences:
                certainty = str(consequence.get("certainty", "possible")).casefold()
                if certainty not in summary:
                    certainty = "possible"
                raw_count = consequence.get(
                    "unique_target_count", consequence.get("count", 0)
                )
                try:
                    count = int(raw_count or 0)
                except (TypeError, ValueError):
                    count = 0
                target_ids = {
                    str(item) for item in consequence.get("target_ids", ()) if item
                }
                for target_id in target_ids:
                    previous = identified_targets.get(target_id)
                    if previous is None or ranks[certainty] > ranks[previous]:
                        identified_targets[target_id] = certainty
                unidentified[certainty] += max(0, count - len(target_ids))
        for certainty in identified_targets.values():
            summary[certainty] += 1
        for certainty, count in unidentified.items():
            summary[certainty] += count
        return summary

    @property
    def root_cause_count(self) -> int:
        return len(self.anomalies)

    @property
    def validation_level(self) -> str:
        if self.errors:
            return "error"
        if self.status in {
            Status.FICHIER_MANQUANT,
            Status.SANS_REFERENCE,
            Status.ERREUR,
        }:
            return "error"
        error_severities = {"error", "critical", "danger", "high", "fatal"}
        if any(
            str(anomaly.severity).casefold() in error_severities
            for anomaly in self.anomalies
        ):
            return "error"
        if self.anomalies or self.warnings:
            return "warning"
        return "ok"

    def finalize_status(self) -> None:
        if self.status in {Status.FICHIER_MANQUANT, Status.SANS_REFERENCE, Status.ERREUR}:
            return
        self.status = Status.ANOMALIES if self.total_anomalies else Status.CONFORME

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["status"] = self.status.value
        data["total_anomalies"] = self.total_anomalies
        data["root_cause_count"] = self.root_cause_count
        data["validation_level"] = self.validation_level
        data["counts"] = self.counts
        data["counts_by_code"] = self.counts_by_code
        data["root_counts_by_code"] = self.root_counts_by_code
        data["impact_summary"] = self.impact_summary
        return data


@dataclass(slots=True)
class FileCandidate:
    key: str
    country: str
    path: Path


@dataclass(slots=True)
class RunSummary:
    results: list[CountryResult]
    sent_dir: str
    received_dir: str
    generated_at: str
    duration_seconds: float
    version: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "sent_dir": self.sent_dir,
            "received_dir": self.received_dir,
            "generated_at": self.generated_at,
            "duration_seconds": self.duration_seconds,
            "version": self.version,
            "results": [result.to_dict() for result in self.results],
        }
