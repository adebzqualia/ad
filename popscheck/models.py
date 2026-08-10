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
        counts = {"feuilles": 0, "colonnes": 0, "lignes": 0, "fichier": 0}
        for anomaly in self.anomalies:
            counts[anomaly.category] = counts.get(anomaly.category, 0) + max(0, anomaly.impact)
        return counts

    def finalize_status(self) -> None:
        if self.status in {Status.FICHIER_MANQUANT, Status.SANS_REFERENCE, Status.ERREUR}:
            return
        self.status = Status.ANOMALIES if self.total_anomalies else Status.CONFORME

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["status"] = self.status.value
        data["total_anomalies"] = self.total_anomalies
        data["counts"] = self.counts
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

