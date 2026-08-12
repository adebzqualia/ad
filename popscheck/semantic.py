from __future__ import annotations

import hashlib
import re
from collections import defaultdict, deque
from typing import Any, Iterable, Mapping

from openpyxl.utils.cell import column_index_from_string

from .config import AppConfig
from .formulas import (
    CellAddress,
    FormulaChange,
    FormulaChangeKind,
    FormulaInventory,
    MappingDirection,
    ParsedFormula,
    SheetCoordinateMapping,
    build_dependency_graph,
    build_inter_sheet_dependencies,
    compare_formula_inventories,
    extract_formulas,
    parse_formula,
)
from .models import Anomaly, CountryResult
from .workbook import WorkbookSnapshot


_FORMULA_CODES = {
    FormulaChangeKind.ADDED: "FORMULA_ADDED",
    FormulaChangeKind.REMOVED: "FORMULA_REMOVED",
    FormulaChangeKind.REPLACED: "FORMULA_REPLACED_BY_VALUE",
    FormulaChangeKind.MODIFIED: "FORMULA_LOGIC_CHANGED",
    FormulaChangeKind.BROKEN_REFERENCE: "INVALID_REFERENCE",
}

_FORMULA_MESSAGES = {
    FormulaChangeKind.ADDED: "Formule ajoutée dans une cellule qui n'en contenait pas",
    FormulaChangeKind.REMOVED: "Formule attendue supprimée",
    FormulaChangeKind.REPLACED: "Formule remplacée par une valeur fixe",
    FormulaChangeKind.MODIFIED: "Logique de formule modifiée",
    FormulaChangeKind.BROKEN_REFERENCE: "Référence Excel invalide (#REF!)",
}

_CERTAINTY_RANK = {"possible": 0, "probable": 1, "confirmed": 2}
_DYNAMIC_REFERENCE_FUNCTIONS = ("INDIRECT(", "OFFSET(", "ADDRESS(")
_OPAQUE_NAME_RE = re.compile(r"^[A-Z_\\][A-Z0-9_.\\]*$", re.IGNORECASE)


def _sheet_reference(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def _qualified(sheet_name: str, coordinate: str) -> str:
    return f"{_sheet_reference(sheet_name)}!{coordinate}"


def _mapping_for(
    sheet_name: str,
    mappings: Mapping[str, SheetCoordinateMapping],
) -> SheetCoordinateMapping | None:
    direct = mappings.get(sheet_name)
    if direct is not None:
        return direct
    folded = sheet_name.casefold()
    return next(
        (mapping for name, mapping in mappings.items() if name.casefold() == folded),
        None,
    )


def _canonical_address(
    change: FormulaChange,
    mappings: Mapping[str, SheetCoordinateMapping],
) -> CellAddress | None:
    if change.expected_address is not None:
        mapping = _mapping_for(change.expected_address.sheet, mappings)
        if mapping is None:
            return None
        if (
            mapping.expected_row(change.expected_address.row) is None
            or mapping.expected_column(change.expected_address.column) is None
        ):
            # The host belongs to a removed axis; the structural anomaly owns it.
            return None
        return change.expected_address
    if change.observed_address is None:
        return None
    mapping = _mapping_for(change.observed_address.sheet, mappings)
    if mapping is None:
        return None
    row = mapping.observed_row(change.observed_address.row)
    column = mapping.observed_column(change.observed_address.column)
    if row is None or column is None:
        # Formula located in an added row/column: content is part of that root.
        return None
    return CellAddress(change.observed_address.sheet, row, column)


def _observed_address(
    canonical: CellAddress,
    mappings: Mapping[str, SheetCoordinateMapping],
) -> CellAddress | None:
    mapping = _mapping_for(canonical.sheet, mappings)
    if mapping is None:
        return None
    row = mapping.expected_row(canonical.row)
    column = mapping.expected_column(canonical.column)
    if row is None or column is None:
        return None
    return CellAddress(canonical.sheet, row, column)


def _formula_policy_allows(
    address: CellAddress,
    config: AppConfig,
) -> bool:
    rule = config.rule_for(address.sheet)
    return (
        not rule.ignore
        and rule.cell_is_monitored(address.row, address.column)
        and not rule.formula_is_allowed(address.row, address.column)
    )


def _canonical_observed_host(
    address: CellAddress,
    mappings: Mapping[str, SheetCoordinateMapping],
) -> CellAddress | None:
    mapping = _mapping_for(address.sheet, mappings)
    if mapping is None:
        return None
    row = mapping.observed_row(address.row)
    column = mapping.observed_column(address.column)
    if row is None or column is None:
        return None
    return CellAddress(address.sheet, row, column)


def _received_only_structural_root(
    address: CellAddress,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
) -> Anomaly | None:
    mapping = _mapping_for(address.sheet, mappings)
    missing_row = mapping is not None and mapping.observed_row(address.row) is None
    missing_column = (
        mapping is not None and mapping.observed_column(address.column) is None
    )
    candidates: list[Anomaly] = []
    for anomaly in result.anomalies:
        if (anomaly.sheet or "").casefold() != address.sheet.casefold():
            continue
        if anomaly.code == "SHEET_ADDED" and mapping is None:
            candidates.append(anomaly)
            continue
        if anomaly.code == "ROW_ADDED" and missing_row:
            span = _parse_axis_span(anomaly.observed_position, column=False)
            if span and span[0] <= address.row <= span[1]:
                candidates.append(anomaly)
        elif anomaly.code == "COLUMN_ADDED" and missing_column:
            span = _parse_axis_span(anomaly.observed_position, column=True)
            if span and span[0] <= address.column <= span[1]:
                candidates.append(anomaly)
    if not candidates:
        return None
    # One received cell can sit at the intersection of an added row and column.
    # Pick one deterministic owner so global impact counts do not double-count it.
    candidates.sort(
        key=lambda item: (
            {"SHEET_ADDED": 0, "ROW_ADDED": 1, "COLUMN_ADDED": 2}.get(
                item.code, 9
            ),
            str(item.observed_position or ""),
        )
    )
    return candidates[0]


def _attach_received_only_formula_impacts(
    observed_inventory: FormulaInventory,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
    config: AppConfig,
    *,
    compare_formulas: bool,
    compare_dependencies: bool,
) -> None:
    for address, parsed in observed_inventory.formulas.items():
        if _canonical_observed_host(address, mappings) is not None:
            continue
        root = _received_only_structural_root(address, result, mappings)
        if root is None:
            continue
        if compare_formulas:
            _attach_consequence(
                root,
                code="FORMULA_ADDED",
                certainty="confirmed",
                count=1,
                affected_sheets=[address.sheet],
                sample_locations=[address.qualified],
                explanation="La zone structurelle ajoutée contient une nouvelle formule.",
                max_samples=config.analysis.max_dependency_samples,
            )
        if parsed.has_ref_error:
            _attach_consequence(
                root,
                code="INVALID_REFERENCE",
                certainty="confirmed",
                count=1,
                affected_sheets=[address.sheet],
                sample_locations=[address.qualified],
                explanation="La nouvelle formule contient #REF!.",
                max_samples=config.analysis.max_dependency_samples,
            )
        if not compare_dependencies:
            continue
        external_books = _external_books(parsed)
        if external_books:
            _attach_consequence(
                root,
                code="EXTERNAL_LINK_ADDED",
                certainty="confirmed",
                count=len(external_books),
                affected_sheets=[address.sheet],
                sample_locations=[address.qualified],
                explanation=(
                    "La zone ajoutée introduit un lien vers un classeur externe : "
                    + ", ".join(sorted(external_books))
                    + "."
                ),
                max_samples=config.analysis.max_dependency_samples,
            )
        internal_targets = {
            reference.sheet
            for reference in parsed.references
            if reference.explicit_sheet
            and reference.sheet
            and not reference.is_external
            and not reference.is_broken
        }
        if internal_targets:
            _attach_consequence(
                root,
                code="DEPENDENCY_ADDED",
                certainty="confirmed",
                count=len(internal_targets),
                affected_sheets=[address.sheet],
                sample_locations=[address.qualified],
                explanation=(
                    "La nouvelle formule ajoute une dépendance vers : "
                    + ", ".join(sorted(internal_targets, key=str.casefold))
                    + "."
                ),
                max_samples=config.analysis.max_dependency_samples,
            )


def _dependency_inventory_for_policy(
    inventory: FormulaInventory,
    mappings: Mapping[str, SheetCoordinateMapping],
    config: AppConfig,
    *,
    observed_side: bool,
) -> FormulaInventory:
    """Keep dependency sources that are inside the canonical audit policy."""

    formulas: dict[CellAddress, ParsedFormula] = {}
    for address, parsed in inventory.formulas.items():
        if observed_side:
            canonical = _canonical_observed_host(address, mappings)
        else:
            mapping = _mapping_for(address.sheet, mappings)
            canonical = (
                address
                if mapping is not None
                and mapping.expected_row(address.row) is not None
                and mapping.expected_column(address.column) is not None
                else None
            )
        if canonical is not None and _formula_policy_allows(canonical, config):
            formulas[address] = parsed
    return FormulaInventory(
        formulas=formulas,
        literal_cells=frozenset(),
        sheet_names=inventory.sheet_names,
        defined_names=inventory.defined_names,
        table_names=inventory.table_names,
        defined_name_definitions=dict(inventory.defined_name_definitions),
        table_definitions=dict(inventory.table_definitions),
        warnings=list(inventory.warnings),
    )


def _formula_severity(
    change: FormulaChange,
    canonical: CellAddress,
    config: AppConfig,
) -> str:
    critical = config.rule_for(canonical.sheet).cell_is_critical(
        canonical.row, canonical.column
    )
    if change.kind in {
        FormulaChangeKind.REMOVED,
        FormulaChangeKind.REPLACED,
        FormulaChangeKind.BROKEN_REFERENCE,
    }:
        return "error"
    if critical:
        return "error"
    return "warning"


def _display_received_value(
    observed: WorkbookSnapshot,
    address: CellAddress | None,
) -> str:
    if address is None or address.sheet not in observed.sheets:
        return "Absent"
    feature = observed.sheets[address.sheet].cells.get((address.row, address.column))
    if feature is None or feature.value is None:
        return "Absent"
    rendered = str(feature.value)
    return rendered if len(rendered) <= 240 else rendered[:239] + "…"


def _parse_axis_span(position: Any, *, column: bool) -> tuple[int, int] | None:
    if isinstance(position, int) and not isinstance(position, bool):
        return position, position
    text = str(position or "").strip()
    if not text:
        return None
    parts = text.split(":", 1)
    try:
        if column:
            values = [column_index_from_string(part.replace("$", "")) for part in parts]
        else:
            values = [int(part.replace("$", "")) for part in parts]
    except (TypeError, ValueError):
        return None
    if len(values) == 1:
        values.append(values[0])
    return min(values), max(values)


def _overlaps(start: int | None, end: int | None, span: tuple[int, int]) -> bool:
    if start is None and end is None:
        return True
    left = start if start is not None else end
    right = end if end is not None else start
    assert left is not None and right is not None
    low, high = sorted((left, right))
    return not (high < span[0] or span[1] < low)


def _formula_targets_root(formula: ParsedFormula, root: Anomaly) -> bool:
    root_sheet = (root.sheet or "").casefold()
    if not root_sheet:
        return False
    if root.code == "SHEET_REMOVED":
        return any(
            reference.sheet
            and reference.sheet.casefold() == root_sheet
            and not reference.is_external
            for reference in formula.references
        )
    if root.code == "COLUMN_REMOVED":
        span = _parse_axis_span(root.expected_position, column=True)
        if span is None:
            return False
        return any(
            reference.sheet
            and reference.sheet.casefold() == root_sheet
            and not reference.is_external
            and _overlaps(reference.start_column, reference.end_column, span)
            for reference in formula.references
        )
    if root.code == "ROW_REMOVED":
        span = _parse_axis_span(root.expected_position, column=False)
        if span is None:
            return False
        return any(
            reference.sheet
            and reference.sheet.casefold() == root_sheet
            and not reference.is_external
            and _overlaps(reference.start_row, reference.end_row, span)
            for reference in formula.references
        )
    return False


def _insertion_boundary(
    root: Anomaly,
    mappings: Mapping[str, SheetCoordinateMapping],
) -> tuple[int, int] | None:
    is_column = root.code == "COLUMN_ADDED"
    if not is_column and root.code != "ROW_ADDED":
        return None
    observed_span = _parse_axis_span(root.observed_position, column=is_column)
    mapping = _mapping_for(root.sheet or "", mappings)
    if observed_span is None or mapping is None:
        return None
    axis_mapping = mapping.columns if is_column else mapping.rows
    if not axis_mapping:
        return None
    before = [
        (observed, expected)
        for expected, observed in axis_mapping.items()
        if observed < observed_span[0]
    ]
    after = [
        (observed, expected)
        for expected, observed in axis_mapping.items()
        if observed > observed_span[1]
    ]
    if not before or not after:
        return None
    left_expected = max(before)[1]
    right_expected = min(after)[1]
    if right_expected != left_expected + 1:
        return None
    return left_expected, right_expected


def _formula_spans_insertion(
    formula: ParsedFormula,
    root: Anomaly,
    mappings: Mapping[str, SheetCoordinateMapping],
) -> bool:
    boundary = _insertion_boundary(root, mappings)
    root_sheet = (root.sheet or "").casefold()
    if boundary is None or not root_sheet:
        return False
    is_column = root.code == "COLUMN_ADDED"
    for reference in formula.references:
        if (
            not reference.sheet
            or reference.sheet.casefold() != root_sheet
            or reference.is_external
        ):
            continue
        start = reference.start_column if is_column else reference.start_row
        end = reference.end_column if is_column else reference.end_row
        if start is None and end is None:
            # A whole-row range spans every column and a whole-column range
            # spans every row. Opaque references remain uncertain elsewhere.
            if (is_column and reference.kind == "row_range") or (
                not is_column and reference.kind == "column_range"
            ):
                return True
            continue
        low = min(value for value in (start, end) if value is not None)
        high = max(value for value in (start, end) if value is not None)
        if low <= boundary[0] and boundary[1] <= high:
            return True
    return False


def _has_dynamic_reference(formula: ParsedFormula) -> bool:
    normalized = formula.normalized.upper()
    return any(token in normalized for token in _DYNAMIC_REFERENCE_FUNCTIONS)


def _dynamic_target_sheets(formula: ParsedFormula) -> set[str]:
    quoted = {
        match.replace("''", "'").casefold()
        for match in re.findall(r"'((?:[^']|'')+)'!", formula.raw)
    }
    bare = {
        match.casefold()
        for match in re.findall(
            r"(?<![A-Za-z0-9_'])\b([A-Za-z_][A-Za-z0-9_.]*)!",
            formula.raw,
        )
    }
    return quoted | bare


def _three_dimensional_targets_sheet(
    formula: ParsedFormula,
    target_sheet: str,
    sheet_order: Iterable[str],
) -> bool:
    ordered = list(sheet_order)
    positions = {name.casefold(): index for index, name in enumerate(ordered)}
    target_position = positions.get(target_sheet.casefold())
    if target_position is None:
        return False
    for reference in formula.references:
        if reference.kind != "three_dimensional" or not reference.sheet:
            continue
        endpoints = reference.sheet.split(":", 1)
        if len(endpoints) != 2:
            continue
        start = positions.get(endpoints[0].casefold())
        end = positions.get(endpoints[1].casefold())
        if start is None or end is None:
            continue
        if min(start, end) <= target_position <= max(start, end):
            return True
    return False


def _structural_candidates(
    formula: ParsedFormula | None,
    anomalies: Iterable[Anomaly],
) -> list[Anomaly]:
    if formula is None:
        return []
    return [
        anomaly
        for anomaly in anomalies
        if anomaly.code in {"SHEET_REMOVED", "COLUMN_REMOVED", "ROW_REMOVED"}
        and _formula_targets_root(formula, anomaly)
    ]


def _attach_consequence(
    root: Anomaly,
    *,
    code: str,
    certainty: str,
    count: int,
    affected_sheets: Iterable[str],
    sample_locations: Iterable[str],
    explanation: str,
    max_samples: int,
) -> None:
    affected = sorted(set(affected_sheets), key=str.casefold)
    all_samples = list(dict.fromkeys(sample_locations))
    samples = all_samples[:max_samples]
    target_ids = {
        hashlib.blake2b(sample.encode("utf-8"), digest_size=8).hexdigest()
        for sample in all_samples
    }
    if certainty == "confirmed" and target_ids:
        for lower_impact in list(root.consequences):
            lower_certainty = str(
                lower_impact.get("certainty", "possible")
            ).casefold()
            if _CERTAINTY_RANK.get(lower_certainty, 0) >= _CERTAINTY_RANK[certainty]:
                continue
            lower_targets = set(lower_impact.get("target_ids", ()))
            overlap = lower_targets & target_ids
            if not overlap:
                continue
            remaining_targets = lower_targets - overlap
            remaining_count = max(
                0, int(lower_impact.get("count", 0) or 0) - len(overlap)
            )
            if remaining_count == 0:
                root.consequences.remove(lower_impact)
                continue
            lower_impact["target_ids"] = sorted(remaining_targets)
            lower_impact["count"] = remaining_count
            lower_impact["unique_target_count"] = remaining_count
            lower_impact["sample_locations"] = [
                sample
                for sample in lower_impact.get("sample_locations", ())
                if hashlib.blake2b(
                    str(sample).encode("utf-8"), digest_size=8
                ).hexdigest()
                not in overlap
            ][:max_samples]
    existing = next(
        (item for item in root.consequences if item.get("code") == code),
        None,
    )
    if existing is not None:
        existing_certainty = str(existing.get("certainty", "possible"))
        if _CERTAINTY_RANK.get(certainty, 0) > _CERTAINTY_RANK.get(
            existing_certainty, 0
        ):
            existing["certainty"] = certainty
        existing_targets = set(existing.get("target_ids", ()))
        existing_count = max(0, int(existing.get("count", 0) or 0))
        newly_identified = len(target_ids - existing_targets)
        unidentified = max(0, count - len(target_ids))
        merged_count = existing_count + newly_identified + unidentified
        existing["count"] = merged_count
        merged_targets = existing_targets | target_ids
        existing["unique_target_count"] = (
            len(merged_targets) if merged_targets else merged_count
        )
        existing["target_ids"] = sorted(merged_targets)
        existing["affected_sheets"] = sorted(
            set(existing.get("affected_sheets", ())) | set(affected),
            key=str.casefold,
        )
        existing["sample_locations"] = list(
            dict.fromkeys([*existing.get("sample_locations", ()), *samples])
        )[:max_samples]
        return
    digest = hashlib.blake2b(
        (
            code
            + "|"
            + (root.sheet or "workbook")
            + "|"
            + str(root.expected_position or root.observed_position or "")
            + "|"
            + "|".join(affected)
            + "|"
            + "|".join(samples)
        ).encode("utf-8"),
        digest_size=8,
    ).hexdigest()
    root.consequences.append(
        {
            "id": f"impact-{digest}",
            "code": code,
            "certainty": certainty,
            "target_type": "formula",
            "count": count,
            "unique_target_count": len(target_ids) if target_ids else count,
            "target_ids": sorted(target_ids),
            "affected_sheets": affected,
            "sample_locations": samples,
            "explanation": explanation,
            "human_review_required": certainty == "possible",
        }
    )


def _attach_structural_dependency_impacts(
    expected_inventory: FormulaInventory,
    observed_inventory: FormulaInventory,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
    config: AppConfig,
    excluded_formula_hosts: set[CellAddress] | None = None,
) -> None:
    excluded_formula_hosts = excluded_formula_hosts or set()
    structural_roots = [
        anomaly
        for anomaly in result.anomalies
        if anomaly.code
        in {
            "SHEET_REMOVED",
            "COLUMN_REMOVED",
            "ROW_REMOVED",
            "COLUMN_ADDED",
            "ROW_ADDED",
        }
    ]
    if not structural_roots:
        return
    graph = build_dependency_graph(expected_inventory)
    reverse_graph: dict[str, set[str]] = defaultdict(set)
    for source, targets in graph.items():
        for target in targets:
            reverse_graph[target.casefold()].add(source)

    for root in structural_roots:
        confirmed: list[CellAddress] = []
        probable: list[CellAddress] = []
        possible: list[CellAddress] = []
        for address, parsed in expected_inventory.formulas.items():
            if address in excluded_formula_hosts:
                continue
            if address.sheet not in mappings or not _formula_policy_allows(address, config):
                continue
            if root.code.endswith("_ADDED"):
                targets_root = _formula_spans_insertion(
                    parsed, root, mappings
                )
            else:
                targets_root = _formula_targets_root(parsed, root)
            three_dimensional_impact = bool(
                root.code == "SHEET_REMOVED"
                and root.sheet
                and _three_dimensional_targets_sheet(
                    parsed,
                    root.sheet,
                    expected_inventory.sheet_names,
                )
            )
            targets_root = targets_root or three_dimensional_impact
            dynamic_reference = _has_dynamic_reference(parsed) and bool(
                root.sheet
                and (
                    root.sheet.casefold() == address.sheet.casefold()
                    or root.sheet.casefold() in _dynamic_target_sheets(parsed)
                )
            )
            if not targets_root and not dynamic_reference:
                continue
            observed_address = _observed_address(address, mappings)
            if observed_address is None:
                continue
            observed_formula = observed_inventory.formulas.get(observed_address)
            if observed_formula is None:
                # A separate formula root will explain a surviving formula that
                # was removed/replaced. Do not duplicate it as an impact here.
                continue
            if observed_formula.has_ref_error:
                confirmed.append(address)
            elif dynamic_reference or three_dimensional_impact or any(
                reference.kind in {"opaque", "three_dimensional"}
                for reference in parsed.references
                if reference.sheet
                and root.sheet
                and reference.sheet.casefold() == root.sheet.casefold()
            ):
                possible.append(address)
            else:
                probable.append(address)

        if confirmed:
            _attach_consequence(
                root,
                code="INVALID_REFERENCE",
                certainty="confirmed",
                count=len(confirmed),
                affected_sheets=(address.sheet for address in confirmed),
                sample_locations=(address.qualified for address in confirmed),
                explanation="Ces formules contiennent #REF! et dépendaient de l'élément supprimé.",
                max_samples=config.analysis.max_dependency_samples,
            )
        if probable:
            _attach_consequence(
                root,
                code="DEPENDENCY_POTENTIALLY_AFFECTED",
                certainty="probable",
                count=len(probable),
                affected_sheets=(address.sheet for address in probable),
                sample_locations=(address.qualified for address in probable),
                explanation="Ces formules utilisent une plage dont la structure a changé.",
                max_samples=config.analysis.max_dependency_samples,
            )
        if possible:
            _attach_consequence(
                root,
                code="DYNAMIC_OR_OPAQUE_DEPENDENCY",
                certainty="possible",
                count=len(possible),
                affected_sheets=(address.sheet for address in possible),
                sample_locations=(address.qualified for address in possible),
                explanation="La dépendance est opaque ou dynamique et nécessite une vérification humaine.",
                max_samples=config.analysis.max_dependency_samples,
            )

        directly_affected = {address.sheet for address in confirmed + probable + possible}
        queue = deque((sheet, 0) for sheet in directly_affected)
        visited = {sheet.casefold() for sheet in directly_affected}
        downstream: set[str] = set()
        while queue:
            target, depth = queue.popleft()
            if depth >= config.analysis.max_dependency_depth:
                continue
            for source in reverse_graph.get(target.casefold(), ()):
                folded = source.casefold()
                if folded in visited:
                    continue
                visited.add(folded)
                downstream.add(source)
                queue.append((source, depth + 1))
        if downstream:
            _attach_consequence(
                root,
                code="DOWNSTREAM_DEPENDENCY_TO_VERIFY",
                certainty="possible",
                count=len(downstream),
                affected_sheets=downstream,
                sample_locations=(),
                explanation="Ces feuilles consomment en aval une feuille déjà affectée.",
                max_samples=config.analysis.max_dependency_samples,
            )


def _targets(parsed: ParsedFormula | None) -> tuple[set[str], set[str]]:
    internal: set[str] = set()
    external: set[str] = set()
    if parsed is None:
        return internal, external
    for reference in parsed.references:
        if reference.external_book:
            external.add(reference.external_book)
        elif reference.explicit_sheet and reference.sheet:
            internal.add(reference.sheet)
    return internal, external


def _formula_anomaly(
    change: FormulaChange,
    canonical: CellAddress,
    observed: WorkbookSnapshot,
    mappings: Mapping[str, SheetCoordinateMapping],
    config: AppConfig,
) -> Anomaly:
    code = _FORMULA_CODES[change.kind]
    actual = change.observed_address or _observed_address(canonical, mappings)
    expected_value = change.expected_formula or "Aucune formule"
    if change.kind is FormulaChangeKind.REPLACED:
        observed_value = _display_received_value(observed, change.observed_address)
    elif change.observed_formula:
        observed_value = change.observed_formula
    else:
        observed_value = "Absent"
    return Anomaly(
        category="formules",
        code=code,
        message=(
            f"{_FORMULA_MESSAGES[change.kind]} dans la feuille « {canonical.sheet} » "
            f"à l'adresse {canonical.coordinate}."
        ),
        sheet=canonical.sheet,
        element=f"Formule {canonical.coordinate}",
        expected_position=canonical.coordinate,
        observed_position=(actual.coordinate if actual else None),
        severity=_formula_severity(change, canonical, config),
        location=canonical.qualified,
        expected=expected_value,
        observed=observed_value,
        confidence={
            "level": "high",
            "score": 1.0,
            "reasons": ["Cellule hôte réalignée avant comparaison sémantique"],
        },
        details={
            "expected_normalized": change.expected_normalized,
            "observed_normalized": change.observed_normalized,
            "formula_change_kind": change.kind.value,
        },
    )


def _dependency_anomaly(
    *,
    code: str,
    source_sheet: str,
    target_expected: str | None,
    target_observed: str | None,
    locations: Iterable[str],
    severity: str = "warning",
) -> Anomaly:
    samples = list(dict.fromkeys(locations))
    messages = {
        "DEPENDENCY_REMOVED": "Dépendance inter-feuilles supprimée",
        "DEPENDENCY_ADDED": "Dépendance inter-feuilles ajoutée",
        "DEPENDENCY_CHANGED": "Cible d'une dépendance modifiée",
        "MISSING_REFERENCED_OBJECT": "Objet Excel référencé inexistant",
        "EXTERNAL_LINK_ADDED": "Lien externe ajouté",
        "EXTERNAL_LINK_REMOVED": "Lien externe supprimé",
        "EXTERNAL_LINK_CHANGED": "Cible d'un lien externe modifiée",
    }
    return Anomaly(
        category="dependances",
        code=code,
        message=f"{messages[code]} depuis la feuille « {source_sheet} ».",
        sheet=source_sheet,
        element=target_expected or target_observed or "Dépendance",
        expected_position=target_expected,
        observed_position=target_observed,
        severity=severity,
        impact=max(1, len(samples)),
        location=samples[0] if samples else source_sheet,
        expected=target_expected or "Absent",
        observed=target_observed or "Absent",
        confidence={"level": "high", "score": 1.0},
        details={"sample_locations": samples},
    )


def _dependency_reference_labels(parsed: ParsedFormula | None) -> set[str]:
    if parsed is None:
        return set()
    return {
        reference.normalized
        for reference in parsed.references
        if reference.explicit_sheet
        and not reference.is_external
        and not reference.is_broken
    }


def _short_reference_list(references: set[str]) -> str | None:
    if not references:
        return None
    ordered = sorted(references, key=str.casefold)
    shown = ordered[:5]
    suffix = f" (+{len(ordered) - len(shown)})" if len(ordered) > len(shown) else ""
    return ", ".join(shown) + suffix


def _attach_formula_dependency_differences(
    expected_inventory: FormulaInventory,
    observed_inventory: FormulaInventory,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
    formula_roots: Mapping[tuple[str, int, int], Anomaly],
    config: AppConfig,
) -> None:
    observed_lookup = {
        (address.sheet.casefold(), address.row, address.column): (address, parsed)
        for address, parsed in observed_inventory.formulas.items()
    }
    matched_observed: set[CellAddress] = set()

    def report(
        canonical: CellAddress,
        expected_formula: ParsedFormula | None,
        before: set[str],
        after: set[str],
    ) -> None:
        if before == after:
            return
        code = (
            "DEPENDENCY_CHANGED"
            if before and after
            else "DEPENDENCY_REMOVED" if before else "DEPENDENCY_ADDED"
        )
        expected_display = _short_reference_list(before)
        observed_display = _short_reference_list(after)
        changed_count = max(1, len(before ^ after))
        explanation = {
            "DEPENDENCY_CHANGED": "La cible ou la couverture d'une référence a changé.",
            "DEPENDENCY_REMOVED": "Une ou plusieurs références attendues ont été supprimées.",
            "DEPENDENCY_ADDED": "Une ou plusieurs nouvelles références ont été ajoutées.",
        }[code]
        candidates = _structural_candidates(expected_formula, result.anomalies)
        root = formula_roots.get(
            (canonical.sheet.casefold(), canonical.row, canonical.column)
        )
        roots = candidates or ([root] if root is not None else [])
        if roots:
            for owner in roots:
                _attach_consequence(
                    owner,
                    code=code,
                    certainty="confirmed",
                    count=changed_count,
                    affected_sheets=[canonical.sheet],
                    sample_locations=[canonical.qualified],
                    explanation=explanation,
                    max_samples=config.analysis.max_dependency_samples,
                )
            return
        result.anomalies.append(
            _dependency_anomaly(
                code=code,
                source_sheet=canonical.sheet,
                target_expected=expected_display,
                target_observed=observed_display,
                locations=[canonical.qualified],
                severity=(
                    "error"
                    if config.rule_for(canonical.sheet).cell_is_critical(
                        canonical.row, canonical.column
                    )
                    else "warning"
                ),
            )
        )

    known_sheets = tuple(
        dict.fromkeys(expected_inventory.sheet_names + observed_inventory.sheet_names)
    )
    for expected_address, expected_formula in sorted(expected_inventory.formulas.items()):
        observed_address = _observed_address(expected_address, mappings)
        observed_formula: ParsedFormula | None = None
        if observed_address is not None:
            matched = observed_lookup.get(
                (
                    observed_address.sheet.casefold(),
                    observed_address.row,
                    observed_address.column,
                )
            )
            if matched is not None:
                actual_address, raw_formula = matched
                matched_observed.add(actual_address)
                observed_formula = parse_formula(
                    raw_formula.raw,
                    address=actual_address,
                    mappings=mappings,
                    mapping_direction=MappingDirection.OBSERVED_TO_EXPECTED,
                    known_sheets=known_sheets,
                )
        report(
            expected_address,
            expected_formula,
            _dependency_reference_labels(expected_formula),
            _dependency_reference_labels(observed_formula),
        )

    for observed_address, observed_formula in sorted(observed_inventory.formulas.items()):
        if observed_address in matched_observed:
            continue
        canonical = _canonical_observed_host(observed_address, mappings)
        if canonical is None:
            continue
        canonical_formula = parse_formula(
            observed_formula.raw,
            address=observed_address,
            mappings=mappings,
            mapping_direction=MappingDirection.OBSERVED_TO_EXPECTED,
            known_sheets=known_sheets,
        )
        report(
            canonical,
            None,
            set(),
            _dependency_reference_labels(canonical_formula),
        )


def _attach_missing_dependency_targets(
    expected_inventory: FormulaInventory,
    observed_inventory: FormulaInventory,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
    formula_roots: Mapping[tuple[str, int, int], Anomaly],
    config: AppConfig,
) -> None:
    expected_dependencies = {
        (item.source_sheet.casefold(), item.target_sheet.casefold()): item
        for item in build_inter_sheet_dependencies(expected_inventory)
    }
    removed_sheet_roots = {
        (anomaly.sheet or "").casefold(): anomaly
        for anomaly in result.anomalies
        if anomaly.code == "SHEET_REMOVED"
    }
    for dependency in build_inter_sheet_dependencies(observed_inventory):
        if dependency.target_exists:
            continue
        key = (
            dependency.source_sheet.casefold(),
            dependency.target_sheet.casefold(),
        )
        expected_dependency = expected_dependencies.get(key)
        if expected_dependency is not None and not expected_dependency.target_exists:
            # The broken reference already exists in the reference template;
            # it is not attributable to the received country file.
            continue
        samples = [address.qualified for address in dependency.formula_cells]
        sheet_root = removed_sheet_roots.get(dependency.target_sheet.casefold())
        roots: list[Anomaly] = [sheet_root] if sheet_root is not None else []
        if not roots:
            for address in dependency.formula_cells:
                canonical = _canonical_observed_host(address, mappings)
                if canonical is None:
                    continue
                root = formula_roots.get(
                    (canonical.sheet.casefold(), canonical.row, canonical.column)
                )
                if root is not None and all(root is not item for item in roots):
                    roots.append(root)
        if roots:
            for root in roots:
                _attach_consequence(
                    root,
                    code="MISSING_REFERENCED_OBJECT",
                    certainty="confirmed",
                    count=max(1, dependency.reference_count),
                    affected_sheets=[dependency.source_sheet],
                    sample_locations=samples,
                    explanation=(
                        f"La feuille référencée « {dependency.target_sheet} » "
                        "n'existe plus."
                    ),
                    max_samples=config.analysis.max_dependency_samples,
                )
            continue
        result.anomalies.append(
            _dependency_anomaly(
                code="MISSING_REFERENCED_OBJECT",
                source_sheet=dependency.source_sheet,
                target_expected=dependency.target_sheet,
                target_observed=None,
                locations=samples,
                severity="error",
            )
        )


def _opaque_object(reference: Any) -> tuple[str, str] | None:
    if reference.kind != "opaque" or reference.explicit_sheet:
        return None
    body = str(reference.normalized).lstrip("@").strip()
    if "[" in body and body.endswith("]"):
        name = body.split("[", 1)[0]
        return ("table", name) if name else None
    if _OPAQUE_NAME_RE.fullmatch(body):
        return "defined_name", body
    return None


def _object_definition_signature(
    definition: str,
    inventory: FormulaInventory,
    mappings: Mapping[str, SheetCoordinateMapping],
    *,
    observed_side: bool,
) -> str:
    if not definition:
        return ""
    formula = definition if definition.startswith("=") else "=" + definition
    return parse_formula(
        formula,
        mappings=mappings if observed_side else None,
        mapping_direction=(
            MappingDirection.OBSERVED_TO_EXPECTED
            if observed_side
            else MappingDirection.NONE
        ),
        known_sheets=inventory.sheet_names,
    ).normalized


def _object_definition_external_books(
    definition: str,
    inventory: FormulaInventory,
) -> set[str]:
    if not definition:
        return set()
    formula = definition if definition.startswith("=") else "=" + definition
    parsed = parse_formula(formula, known_sheets=inventory.sheet_names)
    return {
        reference.external_book
        for reference in parsed.references
        if reference.external_book
    }


def _attach_missing_named_objects(
    expected_inventory: FormulaInventory,
    observed_inventory: FormulaInventory,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
    formula_roots: Mapping[tuple[str, int, int], Anomaly],
    config: AppConfig,
) -> None:
    expected_names = {name.casefold(): name for name in expected_inventory.defined_names}
    observed_names = {name.casefold() for name in observed_inventory.defined_names}
    expected_tables = {name.casefold(): name for name in expected_inventory.table_names}
    observed_tables = {name.casefold() for name in observed_inventory.table_names}
    expected_name_definitions = {
        name.casefold(): value
        for name, value in expected_inventory.defined_name_definitions.items()
    }
    observed_name_definitions = {
        name.casefold(): value
        for name, value in observed_inventory.defined_name_definitions.items()
    }
    expected_table_definitions = {
        name.casefold(): value
        for name, value in expected_inventory.table_definitions.items()
    }
    observed_table_definitions = {
        name.casefold(): value
        for name, value in observed_inventory.table_definitions.items()
    }
    expected_formula_lookup = {
        (address.sheet.casefold(), address.row, address.column): parsed
        for address, parsed in expected_inventory.formulas.items()
    }
    for observed_address, observed_formula in observed_inventory.formulas.items():
        canonical = _canonical_observed_host(observed_address, mappings)
        if canonical is None:
            continue
        expected_formula = expected_formula_lookup.get(
            (canonical.sheet.casefold(), canonical.row, canonical.column)
        )
        if expected_formula is None:
            continue
        expected_objects = {
            item
            for reference in expected_formula.references
            if (item := _opaque_object(reference)) is not None
        }
        observed_objects = {
            item
            for reference in observed_formula.references
            if (item := _opaque_object(reference)) is not None
        }
        missing: list[tuple[str, str]] = []
        changed: list[tuple[str, str, str, str]] = []
        introduced_external_links: list[tuple[str, set[str]]] = []
        for kind, raw_name in expected_objects & observed_objects:
            folded = raw_name.casefold()
            if kind == "table" and folded in expected_tables and folded not in observed_tables:
                missing.append((kind, expected_tables[folded]))
            elif (
                kind == "defined_name"
                and folded in expected_names
                and folded not in observed_names
            ):
                missing.append((kind, expected_names[folded]))
            else:
                if kind == "table":
                    expected_definition = expected_table_definitions.get(folded, "")
                    observed_definition = observed_table_definitions.get(folded, "")
                    display_name = expected_tables.get(folded, raw_name)
                else:
                    expected_definition = expected_name_definitions.get(folded, "")
                    observed_definition = observed_name_definitions.get(folded, "")
                    display_name = expected_names.get(folded, raw_name)
                if expected_definition and observed_definition:
                    expected_signature = _object_definition_signature(
                        expected_definition,
                        expected_inventory,
                        mappings,
                        observed_side=False,
                    )
                    observed_signature = _object_definition_signature(
                        observed_definition,
                        observed_inventory,
                        mappings,
                        observed_side=True,
                    )
                    if expected_signature != observed_signature:
                        changed.append(
                            (
                                kind,
                                display_name,
                                expected_definition,
                                observed_definition,
                            )
                        )
                        if kind == "defined_name":
                            expected_books = _object_definition_external_books(
                                expected_definition,
                                expected_inventory,
                            )
                            observed_books = _object_definition_external_books(
                                observed_definition,
                                observed_inventory,
                            )
                            if not expected_books and observed_books:
                                introduced_external_links.append(
                                    (display_name, observed_books)
                                )
        if not missing and not changed:
            continue
        missing_labels = [
            ("table " if kind == "table" else "nom défini ") + f"« {name} »"
            for kind, name in sorted(missing, key=lambda item: item[1].casefold())
        ]
        root = formula_roots.get(
            (canonical.sheet.casefold(), canonical.row, canonical.column)
        )

        def attach_introduced_external_links(owner: Anomaly) -> None:
            if not introduced_external_links:
                return
            names = sorted(
                {name for name, _books in introduced_external_links},
                key=str.casefold,
            )
            books = sorted(
                {
                    book
                    for _name, linked_books in introduced_external_links
                    for book in linked_books
                },
                key=str.casefold,
            )
            _attach_consequence(
                owner,
                code="EXTERNAL_LINK_ADDED",
                certainty="confirmed",
                count=len(books),
                affected_sheets=[canonical.sheet],
                sample_locations=[canonical.qualified],
                explanation=(
                    "La définition du nom "
                    + ", ".join(f"« {name} »" for name in names)
                    + " introduit un lien vers un classeur externe : "
                    + ", ".join(books)
                    + "."
                ),
                max_samples=config.analysis.max_dependency_samples,
            )

        if root is not None:
            if missing:
                _attach_consequence(
                    root,
                    code="MISSING_REFERENCED_OBJECT",
                    certainty="confirmed",
                    count=len(missing),
                    affected_sheets=[canonical.sheet],
                    sample_locations=[canonical.qualified],
                    explanation=(
                        "Objet(s) référencé(s) absent(s) : "
                        + ", ".join(missing_labels)
                        + "."
                    ),
                    max_samples=config.analysis.max_dependency_samples,
                )
            if changed:
                _attach_consequence(
                    root,
                    code="DEPENDENCY_CHANGED",
                    certainty="confirmed",
                    count=len(changed),
                    affected_sheets=[canonical.sheet],
                    sample_locations=[canonical.qualified],
                    explanation=(
                        "Définition modifiée : "
                        + ", ".join(item[1] for item in changed)
                        + "."
                    ),
                    max_samples=config.analysis.max_dependency_samples,
                )
                attach_introduced_external_links(root)
            continue
        if missing:
            result.anomalies.append(
                _dependency_anomaly(
                    code="MISSING_REFERENCED_OBJECT",
                    source_sheet=canonical.sheet,
                    target_expected=", ".join(missing_labels),
                    target_observed=None,
                    locations=[canonical.qualified],
                    severity="error",
                )
            )
        if changed:
            expected_display = ", ".join(
                f"{name}: {expected_definition}"
                for _kind, name, expected_definition, _observed_definition in changed
            )
            observed_display = ", ".join(
                f"{name}: {observed_definition}"
                for _kind, name, _expected_definition, observed_definition in changed
            )
            changed_anomaly = _dependency_anomaly(
                code="DEPENDENCY_CHANGED",
                source_sheet=canonical.sheet,
                target_expected=expected_display,
                target_observed=observed_display,
                locations=[canonical.qualified],
                severity=(
                    "error"
                    if config.rule_for(canonical.sheet).cell_is_critical(
                        canonical.row, canonical.column
                    )
                    else "warning"
                ),
            )
            attach_introduced_external_links(changed_anomaly)
            result.anomalies.append(changed_anomaly)


def _external_books(parsed: ParsedFormula | None) -> set[str]:
    if parsed is None:
        return set()
    return {
        reference.external_book
        for reference in parsed.references
        if reference.external_book
    }


def _attach_external_link_differences(
    expected_inventory: FormulaInventory,
    observed_inventory: FormulaInventory,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
    formula_roots: Mapping[tuple[str, int, int], Anomaly],
    config: AppConfig,
) -> None:
    for expected_address, expected_formula in expected_inventory.formulas.items():
        if expected_address.sheet not in mappings or not _formula_policy_allows(
            expected_address, config
        ):
            continue
        observed_address = _observed_address(expected_address, mappings)
        observed_formula = (
            observed_inventory.formulas.get(observed_address)
            if observed_address is not None
            else None
        )
        before = _external_books(expected_formula)
        after = _external_books(observed_formula)
        if before == after:
            continue
        if before and after:
            code = "EXTERNAL_LINK_CHANGED"
        elif before:
            code = "EXTERNAL_LINK_REMOVED"
        else:
            code = "EXTERNAL_LINK_ADDED"
        root = formula_roots.get(
            (expected_address.sheet.casefold(), expected_address.row, expected_address.column)
        )
        if root is not None:
            _attach_consequence(
                root,
                code=code,
                certainty="confirmed",
                count=1,
                affected_sheets=[expected_address.sheet],
                sample_locations=[expected_address.qualified],
                explanation="La cible du classeur externe utilisée par cette formule a changé.",
                max_samples=config.analysis.max_dependency_samples,
            )
        else:
            result.anomalies.append(
                _dependency_anomaly(
                    code=code,
                    source_sheet=expected_address.sheet,
                    target_expected=", ".join(sorted(before)) or None,
                    target_observed=", ".join(sorted(after)) or None,
                    locations=[expected_address.qualified],
                )
            )

    # A newly added formula has no expected host formula, so it is absent from
    # the loop above. Its external workbook is still an operational dependency.
    for observed_address, observed_formula in observed_inventory.formulas.items():
        canonical = _canonical_observed_host(observed_address, mappings)
        if canonical is None or canonical in expected_inventory.formulas:
            continue
        after = _external_books(observed_formula)
        if not after:
            continue
        root = formula_roots.get(
            (canonical.sheet.casefold(), canonical.row, canonical.column)
        )
        if root is not None:
            _attach_consequence(
                root,
                code="EXTERNAL_LINK_ADDED",
                certainty="confirmed",
                count=len(after),
                affected_sheets=[canonical.sheet],
                sample_locations=[canonical.qualified],
                explanation="La nouvelle formule ajoute un lien vers un classeur externe.",
                max_samples=config.analysis.max_dependency_samples,
            )
        else:
            result.anomalies.append(
                _dependency_anomaly(
                    code="EXTERNAL_LINK_ADDED",
                    source_sheet=canonical.sheet,
                    target_expected=None,
                    target_observed=", ".join(sorted(after)),
                    locations=[canonical.qualified],
                )
            )


def compare_formula_layer(
    expected: WorkbookSnapshot,
    observed: WorkbookSnapshot,
    result: CountryResult,
    mappings: Mapping[str, SheetCoordinateMapping],
    config: AppConfig,
) -> None:
    compare_formulas_enabled = config.analysis.compare_formulas
    compare_dependencies_enabled = config.analysis.compare_dependencies
    if not (compare_formulas_enabled or compare_dependencies_enabled):
        return
    try:
        expected_inventory = extract_formulas(expected.path)
        observed_inventory = extract_formulas(observed.path)
        changes = (
            compare_formula_inventories(
                expected_inventory,
                observed_inventory,
                mappings=mappings,
            )
            if compare_formulas_enabled
            else ()
        )
    except Exception as exc:
        result.warnings.append(
            "Analyse sémantique des formules incomplète : "
            f"{type(exc).__name__}: {exc}"
        )
        result.metadata["formula_analysis"] = {
            "status": "partial",
            "error": f"{type(exc).__name__}: {exc}",
        }
        return

    expected_dependency_inventory = _dependency_inventory_for_policy(
        expected_inventory,
        mappings,
        config,
        observed_side=False,
    )
    observed_dependency_inventory = _dependency_inventory_for_policy(
        observed_inventory,
        mappings,
        config,
        observed_side=True,
    )

    _attach_received_only_formula_impacts(
        observed_inventory,
        result,
        mappings,
        config,
        compare_formulas=compare_formulas_enabled,
        compare_dependencies=compare_dependencies_enabled,
    )

    independent_formula_hosts = {
        canonical
        for change in changes
        if change.kind is FormulaChangeKind.MODIFIED
        and (canonical := _canonical_address(change, mappings)) is not None
    }

    if compare_dependencies_enabled:
        _attach_structural_dependency_impacts(
            expected_dependency_inventory,
            observed_dependency_inventory,
            result,
            mappings,
            config,
            excluded_formula_hosts=independent_formula_hosts,
        )

    formula_roots: dict[tuple[str, int, int], Anomaly] = {}
    for change in changes:
        canonical = _canonical_address(change, mappings)
        if canonical is None or not _formula_policy_allows(canonical, config):
            continue
        expected_formula = expected_inventory.formulas.get(canonical)
        candidates = _structural_candidates(expected_formula, result.anomalies)
        observed_parsed = (
            observed_inventory.formulas.get(change.observed_address)
            if change.observed_address is not None
            else None
        )
        if change.kind is FormulaChangeKind.MODIFIED and candidates and (
            observed_parsed and observed_parsed.has_ref_error
        ):
            # The dedicated BROKEN_REFERENCE finding below will be attached to
            # the structural root; keeping MODIFIED would double-count it.
            continue
        key = (canonical.sheet.casefold(), canonical.row, canonical.column)
        if change.kind is FormulaChangeKind.BROKEN_REFERENCE:
            if candidates:
                for root in candidates:
                    _attach_consequence(
                        root,
                        code="INVALID_REFERENCE",
                        certainty="confirmed",
                        count=1,
                        affected_sheets=[canonical.sheet],
                        sample_locations=[canonical.qualified],
                        explanation="Cette formule contient #REF! après la modification structurelle.",
                        max_samples=config.analysis.max_dependency_samples,
                    )
                continue
            existing = formula_roots.get(key)
            if existing is not None:
                existing.severity = "error"
                _attach_consequence(
                    existing,
                    code="INVALID_REFERENCE",
                    certainty="confirmed",
                    count=1,
                    affected_sheets=[canonical.sheet],
                    sample_locations=[canonical.qualified],
                    explanation="La formule modifiée contient #REF!.",
                    max_samples=config.analysis.max_dependency_samples,
                )
                continue

        anomaly = _formula_anomaly(change, canonical, observed, mappings, config)
        result.anomalies.append(anomaly)
        formula_roots[key] = anomaly

    if compare_dependencies_enabled:
        _attach_formula_dependency_differences(
            expected_dependency_inventory,
            observed_dependency_inventory,
            result,
            mappings,
            formula_roots,
            config,
        )
        _attach_missing_dependency_targets(
            expected_dependency_inventory,
            observed_dependency_inventory,
            result,
            mappings,
            formula_roots,
            config,
        )
        _attach_missing_named_objects(
            expected_dependency_inventory,
            observed_dependency_inventory,
            result,
            mappings,
            formula_roots,
            config,
        )
        _attach_external_link_differences(
            expected_dependency_inventory,
            observed_dependency_inventory,
            result,
            mappings,
            formula_roots,
            config,
        )

    inventory_warnings = list(
        dict.fromkeys(expected_inventory.warnings + observed_inventory.warnings)
    )
    if inventory_warnings:
        visible_warnings = inventory_warnings[: config.analysis.max_dependency_samples]
        result.warnings.extend(
            f"Analyse de formule partielle — {warning}"
            for warning in visible_warnings
        )
        hidden_count = len(inventory_warnings) - len(visible_warnings)
        if hidden_count:
            result.warnings.append(
                f"Analyse de formule partielle — {hidden_count} autre(s) "
                "limitation(s) technique(s) sont détaillées dans le JSON."
            )

    result.metadata["formula_analysis"] = {
        "status": "partial" if inventory_warnings else "complete",
        "formula_comparison": "enabled" if compare_formulas_enabled else "disabled",
        "dependency_comparison": (
            "enabled" if compare_dependencies_enabled else "disabled"
        ),
        "expected_formula_count": len(expected_dependency_inventory.formulas),
        "observed_formula_count": len(observed_dependency_inventory.formulas),
        "change_count": sum(
            1
            for change in changes
            if (canonical := _canonical_address(change, mappings)) is not None
            and _formula_policy_allows(canonical, config)
        ),
        "raw_expected_formula_count": len(expected_inventory.formulas),
        "raw_observed_formula_count": len(observed_inventory.formulas),
        "raw_change_count": len(changes),
        "dependency_count_expected": len(
            build_inter_sheet_dependencies(expected_dependency_inventory)
        ),
        "dependency_count_observed": len(
            build_inter_sheet_dependencies(observed_dependency_inventory)
        ),
        "warning_count": len(inventory_warnings),
        "warnings": inventory_warnings,
    }


_ACTION_TEMPLATES: dict[str, tuple[str, tuple[str, ...]]] = {
    "SHEET_REMOVED": (
        "Restaurer la feuille depuis le fichier envoyé.",
        ("Copier la feuille attendue depuis le template.", "Vérifier les formules dépendantes."),
    ),
    "SHEET_ADDED": (
        "Confirmer l'utilité de la feuille ajoutée ou la supprimer.",
        ("Identifier l'origine de la feuille.", "La retirer si elle n'est pas autorisée."),
    ),
    "SHEET_ORDER_CHANGED": (
        "Rétablir l'ordre des feuilles du fichier envoyé.",
        ("Comparer l'ordre des onglets avec le template.", "Replacer les feuilles sans modifier leur contenu."),
    ),
    "SHEET_TYPE_CHANGED": (
        "Restaurer le type de feuille attendu depuis le template.",
        ("Remplacer l'objet incompatible.", "Vérifier les formules et graphiques qui le référencent."),
    ),
    "COLUMN_REMOVED": (
        "Restaurer la colonne depuis le fichier envoyé.",
        ("Réinsérer la colonne attendue.", "Reporter les données puis vérifier les formules dépendantes."),
    ),
    "ROW_REMOVED": (
        "Restaurer la ligne depuis le fichier envoyé.",
        ("Réinsérer la ligne attendue.", "Reporter les données puis vérifier les calculs dépendants."),
    ),
    "COLUMN_ADDED": (
        "Supprimer la colonne ou justifier son ajout.",
        ("Comparer la zone avec le template.", "Retirer l'ajout s'il n'est pas autorisé."),
    ),
    "ROW_ADDED": (
        "Supprimer la ligne ou justifier son ajout.",
        ("Comparer la zone avec le template.", "Retirer l'ajout s'il n'est pas autorisé."),
    ),
    "COLUMN_MOVED": ("Remettre la colonne à sa position attendue.", ("Restaurer l'ordre du template.",)),
    "ROW_MOVED": ("Remettre la ligne à sa position attendue.", ("Restaurer l'ordre du template.",)),
    "FORMULA_REMOVED": ("Restaurer la formule du template.", ("Recopier la formule attendue.", "Vérifier le résultat et les cellules aval.")),
    "FORMULA_REPLACED_BY_VALUE": ("Remplacer la valeur fixe par la formule attendue.", ("Restaurer la formule.", "Vérifier que la cellule se recalcule.")),
    "FORMULA_ADDED": ("Retirer la formule ou faire valider la nouvelle logique.", ("Comparer avec le template.",)),
    "FORMULA_LOGIC_CHANGED": ("Restaurer la formule ou faire valider la modification.", ("Comparer les références et opérateurs.", "Contrôler les résultats aval.")),
    "INVALID_REFERENCE": ("Corriger la formule contenant #REF!.", ("Restaurer la référence attendue.", "Recalculer le classeur.")),
    "DEPENDENCY_REMOVED": ("Restaurer la référence attendue entre les feuilles.", ("Comparer la formule au template.", "Vérifier la feuille source après recalcul.")),
    "DEPENDENCY_ADDED": ("Faire valider ou retirer la nouvelle dépendance.", ("Identifier la nouvelle source.", "Confirmer qu'elle est autorisée.")),
    "DEPENDENCY_CHANGED": ("Restaurer ou faire valider la nouvelle cible de la formule.", ("Comparer la cellule ou plage référencée au template.", "Contrôler le résultat aval.")),
    "MISSING_REFERENCED_OBJECT": ("Restaurer l'objet Excel référencé ou corriger la formule.", ("Vérifier la feuille, table ou le nom défini attendu.", "Recalculer le classeur après correction.")),
    "EXTERNAL_LINK_ADDED": ("Supprimer le lien externe ou le faire autoriser explicitement.", ("Identifier le classeur externe.", "Vérifier sa source et sa portabilité.")),
    "EXTERNAL_LINK_REMOVED": ("Restaurer le lien externe attendu ou valider sa suppression.", ("Comparer la cible au template.",)),
    "EXTERNAL_LINK_CHANGED": ("Restaurer ou valider la nouvelle cible externe.", ("Comparer le chemin et le nom du classeur.", "Contrôler la source avant diffusion.")),
    "CONTROLLED_VALUE_CHANGED": ("Restaurer ou justifier la valeur contrôlée.", ("Comparer avec le template.",)),
    "STRUCTURAL_VALUE_REMOVED": ("Restaurer la valeur structurelle attendue.", ("Reprendre la valeur du template.",)),
    "VALUE_ADDED_OUTSIDE_EDITABLE_ZONE": ("Supprimer ou faire valider la saisie hors zone autorisée.", ("Vérifier la politique de saisie.",)),
}


def _stable_anomaly_id(result: CountryResult, anomaly: Anomaly) -> str:
    payload = "|".join(
        (
            result.key,
            anomaly.code,
            anomaly.sheet or "workbook",
            str(
                anomaly.expected_position
                if anomaly.expected_position is not None
                else anomaly.observed_position or ""
            ),
            str(anomaly.location or ""),
            str(anomaly.element or ""),
        )
    )
    digest = hashlib.blake2b(payload.encode("utf-8"), digest_size=10).hexdigest()
    return f"pc-{digest}"


def _normalise_confidence(anomaly: Anomaly) -> None:
    if anomaly.confidence is not None:
        return
    raw = anomaly.details.get("confidence")
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        score = max(0.0, min(1.0, float(raw)))
        level = "high" if score >= 0.85 else "medium" if score >= 0.6 else "low"
    else:
        token = str(raw or "").casefold()
        if token in {"élevée", "elevee", "high"}:
            level, score = "high", 0.95
        elif token in {"limitée", "limitee", "low"}:
            level, score = "low", 0.45
        else:
            level, score = "medium", 0.7
    anomaly.confidence = {"level": level, "score": score}


def finalize_operational_anomalies(result: CountryResult) -> None:
    for anomaly in result.anomalies:
        if anomaly.id is None:
            anomaly.id = _stable_anomaly_id(result, anomaly)
        _normalise_confidence(anomaly)
        if any(
            item.get("certainty") == "confirmed"
            and item.get("code") in {"INVALID_REFERENCE", "MISSING_REFERENCED_OBJECT"}
            for item in anomaly.consequences
        ):
            anomaly.severity = "error"
        if not anomaly.action:
            summary, steps = _ACTION_TEMPLATES.get(
                anomaly.code,
                (
                    "Vérifier cette différence avec le fichier envoyé.",
                    ("Confirmer si la modification est autorisée.",),
                ),
            )
            anomaly.action = {
                "owner": "country",
                "priority": "high" if anomaly.severity == "error" else "normal",
                "summary": summary,
                "steps": list(steps),
                "verification_after_fix": [
                    "Relancer POPS Check et vérifier que cette cause a disparu."
                ],
                "country_message": (
                    f"Dans la feuille {anomaly.sheet or 'concernée'}, "
                    f"{anomaly.message.rstrip('.').casefold()}. {summary}"
                ),
            }
        grouped = sum(
            max(0, int(item.get("count", 0) or 0)) for item in anomaly.consequences
        )
        if grouped and not anomaly.noise_reduction:
            anomaly.noise_reduction = {
                "grouped_count": grouped,
                "count_method": "semantic",
                "unresolved_count": 0,
            }


__all__ = ["compare_formula_layer", "finalize_operational_anomalies"]
