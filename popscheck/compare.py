from __future__ import annotations

import time
import unicodedata
from collections import defaultdict
from pathlib import Path
from typing import Iterable

from openpyxl.utils import get_column_letter

from .align import AxisAlignment, AxisItem, align_axis, build_axis_items
from .config import AppConfig, SheetRule
from .models import Anomaly, CountryResult, FileCandidate, Status
from .workbook import SheetSnapshot, WorkbookSnapshot, load_snapshot


def _normalized_file_key(relative_path: Path, case_sensitive: bool) -> str:
    without_suffix = relative_path.with_suffix("").as_posix()
    normalized = unicodedata.normalize("NFC", without_suffix)
    return normalized if case_sensitive else normalized.casefold()


def discover_files(directory: str | Path, config: AppConfig) -> dict[str, list[FileCandidate]]:
    root = Path(directory)
    if not root.is_dir():
        raise FileNotFoundError(f"Dossier introuvable : {root}")
    iterator: Iterable[Path] = root.rglob("*") if config.analysis.recursive else root.glob("*")
    candidates: dict[str, list[FileCandidate]] = defaultdict(list)
    extensions = set(config.analysis.extensions)
    for path in iterator:
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in extensions:
            continue
        relative = path.relative_to(root)
        key = _normalized_file_key(relative, config.analysis.case_sensitive_names)
        country = relative.with_suffix("").as_posix().replace("/", " / ")
        candidates[key].append(FileCandidate(key=key, country=country, path=path.resolve()))
    for entries in candidates.values():
        entries.sort(key=lambda entry: (entry.path.name.casefold(), str(entry.path)))
    return dict(candidates)


def _observed_extent(
    expected: SheetSnapshot,
    observed: SheetSnapshot,
    axis: str,
    config: AppConfig,
    rule: SheetRule,
) -> tuple[int, int]:
    expected_extent = expected.extent(
        axis, config.analysis, rule, include_literal_values=True
    )
    observed_natural = observed.extent(
        axis, config.analysis, rule, include_literal_values=True
    )
    if config.analysis.detect_value_only_expansion:
        return expected_extent, observed_natural
    observed_structural = observed.extent(
        axis, config.analysis, rule, include_literal_values=False
    )
    # Une valeur isolée hors du template ne suffit pas à inventer une ligne ou
    # colonne. Les styles, formules, tables, validations et dimensions, eux, le peuvent.
    observed_extent = max(min(observed_natural, expected_extent), observed_structural)
    return expected_extent, observed_extent


def _axis_pipeline(
    expected: SheetSnapshot,
    observed: SheetSnapshot,
    rule: SheetRule,
    common_labels: set[str],
    config: AppConfig,
) -> tuple[
    AxisAlignment,
    AxisAlignment,
    list[AxisItem],
    list[AxisItem],
    list[AxisItem],
    list[AxisItem],
]:
    expected_column_extent, observed_column_extent = _observed_extent(
        expected, observed, "column", config, rule
    )
    expected_row_extent, observed_row_extent = _observed_extent(
        expected, observed, "row", config, rule
    )

    expected_columns = build_axis_items(
        expected, "column", expected_column_extent, rule, common_labels
    )
    observed_columns = build_axis_items(
        observed, "column", observed_column_extent, rule, common_labels
    )
    column_alignment = align_axis(expected_columns, observed_columns, config.analysis)

    expected_rows = build_axis_items(
        expected,
        "row",
        expected_row_extent,
        rule,
        common_labels,
        orthogonal_mapping={index: index for index in range(1, expected_column_extent + 1)},
    )
    observed_rows = build_axis_items(
        observed,
        "row",
        observed_row_extent,
        rule,
        common_labels,
        orthogonal_mapping=column_alignment.observed_to_expected,
        observed_side=True,
    )
    row_alignment = align_axis(expected_rows, observed_rows, config.analysis)

    # Une seconde passe stabilise les signatures lorsqu'une ligne et une colonne
    # ont été modifiées dans le même fichier.
    expected_columns = build_axis_items(
        expected,
        "column",
        expected_column_extent,
        rule,
        common_labels,
        orthogonal_mapping={index: index for index in range(1, expected_row_extent + 1)},
    )
    observed_columns = build_axis_items(
        observed,
        "column",
        observed_column_extent,
        rule,
        common_labels,
        orthogonal_mapping=row_alignment.observed_to_expected,
        observed_side=True,
    )
    column_alignment = align_axis(expected_columns, observed_columns, config.analysis)

    expected_rows = build_axis_items(
        expected,
        "row",
        expected_row_extent,
        rule,
        common_labels,
        orthogonal_mapping={index: index for index in range(1, expected_column_extent + 1)},
    )
    observed_rows = build_axis_items(
        observed,
        "row",
        observed_row_extent,
        rule,
        common_labels,
        orthogonal_mapping=column_alignment.observed_to_expected,
        observed_side=True,
    )
    row_alignment = align_axis(expected_rows, observed_rows, config.analysis)
    return (
        row_alignment,
        column_alignment,
        expected_rows,
        observed_rows,
        expected_columns,
        observed_columns,
    )


def _axis_label(item: AxisItem | None, fallback: str) -> str:
    return item.label if item and item.label else fallback


def _add_axis_anomalies(
    result: CountryResult,
    sheet_name: str,
    axis: str,
    alignment: AxisAlignment,
    expected_items: list[AxisItem],
    observed_items: list[AxisItem],
) -> None:
    expected_by_index = {item.index: item for item in expected_items}
    observed_by_index = {item.index: item for item in observed_items}
    is_row = axis == "row"
    category = "lignes" if is_row else "colonnes"
    noun = "Ligne" if is_row else "Colonne"
    code_prefix = "ROW" if is_row else "COLUMN"

    def position(index: int) -> int | str:
        return index if is_row else get_column_letter(index)

    for expected_index in alignment.removed:
        item = expected_by_index.get(expected_index)
        expected_position = position(expected_index)
        element = _axis_label(item, f"{noun} {expected_position}")
        result.anomalies.append(
            Anomaly(
                category=category,
                code=f"{code_prefix}_REMOVED",
                message=f"{noun} supprimée dans la feuille « {sheet_name} » à la position {expected_position}.",
                sheet=sheet_name,
                element=element,
                expected_position=expected_position,
                observed_position=None,
                severity="warning",
                details={"confidence": "élevée", "evidence_weight": item.information if item else 0},
            )
        )
    for observed_index in alignment.added:
        item = observed_by_index.get(observed_index)
        observed_position = position(observed_index)
        element = _axis_label(item, f"{noun} {observed_position}")
        result.anomalies.append(
            Anomaly(
                category=category,
                code=f"{code_prefix}_ADDED",
                message=f"{noun} ajoutée dans la feuille « {sheet_name} » à la position {observed_position}.",
                sheet=sheet_name,
                element=element,
                expected_position=None,
                observed_position=observed_position,
                severity="warning",
                details={"confidence": "élevée", "evidence_weight": item.information if item else 0},
            )
        )
    for expected_index, observed_index, score in alignment.moved:
        item = expected_by_index.get(expected_index)
        expected_position = position(expected_index)
        observed_position = position(observed_index)
        element = _axis_label(item, f"{noun} {expected_position}")
        result.anomalies.append(
            Anomaly(
                category=category,
                code=f"{code_prefix}_MOVED",
                message=(
                    f"{noun} déplacée dans la feuille « {sheet_name} » : "
                    f"position attendue {expected_position}, position reçue {observed_position}."
                ),
                sheet=sheet_name,
                element=element,
                expected_position=expected_position,
                observed_position=observed_position,
                severity="warning",
                details={"confidence": round(score, 3)},
            )
        )
    if alignment.ambiguities:
        for ambiguity in alignment.ambiguities:
            result.warnings.append(f"{sheet_name} — {noun.lower()}s : {ambiguity}")


def _compare_sheet(
    expected: SheetSnapshot,
    observed: SheetSnapshot,
    result: CountryResult,
    config: AppConfig,
) -> None:
    rule = config.rule_for(expected.name)
    if rule.ignore:
        result.metadata.setdefault("ignored_sheets", []).append(expected.name)
        return
    if expected.kind != observed.kind:
        result.anomalies.append(
            Anomaly(
                category="feuilles",
                code="SHEET_TYPE_CHANGED",
                message=f"Le type de la feuille « {expected.name} » a changé.",
                sheet=expected.name,
                element=expected.name,
                expected_position=expected.kind,
                observed_position=observed.kind,
                severity="error",
            )
        )
        return
    if expected.kind != "worksheet":
        return
    if expected.state != observed.state:
        result.warnings.append(
            f"{expected.name} : visibilité modifiée ({expected.state} → {observed.state})."
        )
    common_labels = (
        expected.labels & observed.labels if config.analysis.use_stable_text_anchors else set()
    )
    (
        row_alignment,
        column_alignment,
        expected_rows,
        observed_rows,
        expected_columns,
        observed_columns,
    ) = _axis_pipeline(expected, observed, rule, common_labels, config)
    _add_axis_anomalies(
        result,
        expected.name,
        "column",
        column_alignment,
        expected_columns,
        observed_columns,
    )
    _add_axis_anomalies(
        result,
        expected.name,
        "row",
        row_alignment,
        expected_rows,
        observed_rows,
    )


def compare_snapshots(
    expected: WorkbookSnapshot,
    observed: WorkbookSnapshot,
    country: str,
    key: str,
    config: AppConfig,
) -> CountryResult:
    result = CountryResult(
        key=key,
        country=country,
        reference_path=str(expected.path),
        received_path=str(observed.path),
        sheet_order_expected=list(expected.sheet_order),
        sheet_order_observed=list(observed.sheet_order),
    )
    result.warnings.extend(expected.warnings)
    result.warnings.extend(observed.warnings)

    expected_names = set(expected.sheet_order)
    observed_names = set(observed.sheet_order)
    for sheet_name in expected.sheet_order:
        if sheet_name not in observed_names:
            result.anomalies.append(
                Anomaly(
                    category="feuilles",
                    code="SHEET_REMOVED",
                    message=f"Feuille attendue absente : « {sheet_name} ».",
                    sheet=sheet_name,
                    element=sheet_name,
                    expected_position=expected.sheet_order.index(sheet_name) + 1,
                    severity="error",
                )
            )
    for sheet_name in observed.sheet_order:
        if sheet_name not in expected_names:
            result.anomalies.append(
                Anomaly(
                    category="feuilles",
                    code="SHEET_ADDED",
                    message=f"Feuille supplémentaire : « {sheet_name} ».",
                    sheet=sheet_name,
                    element=sheet_name,
                    observed_position=observed.sheet_order.index(sheet_name) + 1,
                    severity="warning",
                )
            )

    expected_common = [name for name in expected.sheet_order if name in observed_names]
    observed_common = [name for name in observed.sheet_order if name in expected_names]
    if (
        config.analysis.strict_sheet_order
        and expected_common != observed_common
        and len(expected_common) > 1
    ):
        result.anomalies.append(
            Anomaly(
                category="feuilles",
                code="SHEET_ORDER_CHANGED",
                message="L’ordre relatif des feuilles communes a été modifié.",
                element="Ordre des feuilles",
                expected_position=" > ".join(expected_common),
                observed_position=" > ".join(observed_common),
                severity="warning",
                details={"expected": expected_common, "observed": observed_common},
            )
        )

    for sheet_name in expected_common:
        _compare_sheet(
            expected.sheets[sheet_name], observed.sheets[sheet_name], result, config
        )
    result.metadata["analyzed_sheets"] = len(expected_common)
    result.finalize_status()
    return result


def compare_workbooks(
    reference_path: str | Path,
    received_path: str | Path,
    *,
    country: str | None = None,
    key: str | None = None,
    config: AppConfig | None = None,
) -> CountryResult:
    app_config = config or AppConfig()
    reference = Path(reference_path).resolve()
    received = Path(received_path).resolve()
    display_country = country or reference.stem
    result_key = key or unicodedata.normalize("NFC", display_country).casefold()
    started = time.perf_counter()
    try:
        expected_snapshot = load_snapshot(reference, app_config.analysis)
        observed_snapshot = load_snapshot(received, app_config.analysis)
        result = compare_snapshots(
            expected_snapshot, observed_snapshot, display_country, result_key, app_config
        )
    except Exception as exc:
        result = CountryResult(
            key=result_key,
            country=display_country,
            reference_path=str(reference),
            received_path=str(received),
            status=Status.ERREUR,
            errors=[f"{type(exc).__name__}: {exc}"],
        )
    result.metadata["duration_seconds"] = round(time.perf_counter() - started, 3)
    return result


def _collision_result(
    key: str,
    sent: list[FileCandidate],
    received: list[FileCandidate],
) -> CountryResult:
    available = sent or received
    country = available[0].country if available else key
    messages: list[str] = []
    if len(sent) > 1:
        messages.append(
            "Plusieurs fichiers de référence correspondent au même pays : "
            + ", ".join(str(item.path) for item in sent)
        )
    if len(received) > 1:
        messages.append(
            "Plusieurs fichiers reçus correspondent au même pays : "
            + ", ".join(str(item.path) for item in received)
        )
    return CountryResult(
        key=key,
        country=country,
        reference_path=str(sent[0].path) if len(sent) == 1 else None,
        received_path=str(received[0].path) if len(received) == 1 else None,
        status=Status.ERREUR,
        errors=messages,
    )


def analyze_directories(
    sent_dir: str | Path,
    received_dir: str | Path,
    config: AppConfig | None = None,
) -> list[CountryResult]:
    app_config = config or AppConfig()
    sent_files = discover_files(sent_dir, app_config)
    received_files = discover_files(received_dir, app_config)
    results: list[CountryResult] = []
    for key in sorted(set(sent_files) | set(received_files)):
        sent = sent_files.get(key, [])
        received = received_files.get(key, [])
        if len(sent) > 1 or len(received) > 1:
            results.append(_collision_result(key, sent, received))
            continue
        country = (sent or received)[0].country
        if not sent:
            results.append(
                CountryResult(
                    key=key,
                    country=country,
                    received_path=str(received[0].path),
                    status=Status.SANS_REFERENCE,
                    errors=["Aucun fichier de référence correspondant dans le dossier sent."],
                )
            )
            continue
        if not received:
            results.append(
                CountryResult(
                    key=key,
                    country=country,
                    reference_path=str(sent[0].path),
                    status=Status.FICHIER_MANQUANT,
                    errors=["Aucun fichier reçu correspondant dans le dossier received."],
                )
            )
            continue
        results.append(
            compare_workbooks(
                sent[0].path,
                received[0].path,
                country=country,
                key=key,
                config=app_config,
            )
        )
    return results
