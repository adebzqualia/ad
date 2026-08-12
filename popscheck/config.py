from __future__ import annotations

import fnmatch
import math
import tomllib
from dataclasses import dataclass, field, replace
from pathlib import Path

from openpyxl.utils.cell import range_boundaries


SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")
EXCEL_MAX_ROW = 1_048_576
EXCEL_MAX_COLUMN = 16_384


@dataclass(frozen=True, slots=True)
class SheetRule:
    pattern: str = "*"
    ignore: bool = False
    editable_ranges: tuple[str, ...] = ()
    formula_allowed_ranges: tuple[str, ...] = ()
    monitored_ranges: tuple[str, ...] = ()
    controlled_ranges: tuple[str, ...] = ()
    critical_ranges: tuple[str, ...] = ()

    def matches(self, sheet_name: str) -> bool:
        return fnmatch.fnmatchcase(sheet_name.casefold(), self.pattern.casefold())

    def cell_is_editable(self, row: int, column: int) -> bool:
        protected = any(
            _contains(cell_range, row, column)
            for cell_range in self.controlled_ranges + self.critical_ranges
        )
        return not protected and any(
            _contains(cell_range, row, column) for cell_range in self.editable_ranges
        )

    def cell_is_monitored(self, row: int, column: int) -> bool:
        if any(
            _contains(cell_range, row, column)
            for cell_range in self.controlled_ranges + self.critical_ranges
        ):
            return True
        if not self.monitored_ranges:
            return True
        return any(_contains(cell_range, row, column) for cell_range in self.monitored_ranges)

    def formula_is_allowed(self, row: int, column: int) -> bool:
        protected = any(
            _contains(cell_range, row, column)
            for cell_range in self.controlled_ranges + self.critical_ranges
        )
        return not protected and any(
            _contains(cell_range, row, column)
            for cell_range in self.formula_allowed_ranges
        )

    def cell_is_controlled(self, row: int, column: int) -> bool:
        return any(_contains(cell_range, row, column) for cell_range in self.controlled_ranges)

    def cell_is_critical(self, row: int, column: int) -> bool:
        return any(_contains(cell_range, row, column) for cell_range in self.critical_ranges)

    def axis_coordinate_is_monitored(self, axis: str, index: int) -> bool:
        """Check one known axis while the other coordinate may still be unknown."""

        if not self.monitored_ranges:
            return True
        for configured_range in (
            self.monitored_ranges + self.controlled_ranges + self.critical_ranges
        ):
            min_col, min_row, max_col, max_row = _range_bounds(configured_range)
            if axis == "row" and min_row <= index <= max_row:
                return True
            if axis == "column" and min_col <= index <= max_col:
                return True
        return False

    def monitored_extent(self) -> tuple[int, int] | None:
        configured_ranges = (
            self.monitored_ranges + self.controlled_ranges + self.critical_ranges
        )
        if not configured_ranges:
            return None
        max_row = 0
        max_column = 0
        for configured_range in configured_ranges:
            _min_col, _min_row, range_max_col, range_max_row = _range_bounds(
                configured_range
            )
            max_row = max(max_row, range_max_row)
            max_column = max(max_column, range_max_col)
        return max_row, max_column


@dataclass(frozen=True, slots=True)
class AnalysisConfig:
    extensions: tuple[str, ...] = SUPPORTED_EXTENSIONS
    recursive: bool = False
    case_sensitive_names: bool = False
    strict_sheet_order: bool = True
    use_stable_text_anchors: bool = True
    min_axis_similarity: float = 0.62
    move_min_similarity: float = 0.84
    ambiguity_margin: float = 0.08
    max_cells_per_sheet: int = 500_000
    max_rows: int = 50_000
    max_columns: int = 2_000
    max_style_gap: int = 25
    detect_value_only_expansion: bool = False
    report_ambiguities: bool = True
    compare_formulas: bool = True
    compare_dependencies: bool = True
    compare_controlled_values: bool = True
    max_dependency_depth: int = 8
    max_dependency_samples: int = 20
    numeric_absolute_tolerance: float = 0.0
    numeric_relative_tolerance: float = 0.0


@dataclass(frozen=True, slots=True)
class AppConfig:
    analysis: AnalysisConfig = field(default_factory=AnalysisConfig)
    sheet_rules: tuple[SheetRule, ...] = ()

    def rule_for(self, sheet_name: str) -> SheetRule:
        selected = SheetRule()
        for rule in self.sheet_rules:
            if rule.matches(sheet_name):
                selected = replace(
                    selected,
                    pattern=rule.pattern,
                    ignore=rule.ignore,
                    editable_ranges=selected.editable_ranges + rule.editable_ranges,
                    formula_allowed_ranges=(
                        selected.formula_allowed_ranges
                        + rule.formula_allowed_ranges
                    ),
                    monitored_ranges=selected.monitored_ranges + rule.monitored_ranges,
                    controlled_ranges=selected.controlled_ranges + rule.controlled_ranges,
                    critical_ranges=selected.critical_ranges + rule.critical_ranges,
                )
        return selected


def _contains(cell_range: str, row: int, column: int) -> bool:
    min_col, min_row, max_col, max_row = _range_bounds(cell_range)
    return min_row <= row <= max_row and min_col <= column <= max_col


def _range_bounds(cell_range: str) -> tuple[int, int, int, int]:
    try:
        bounds = range_boundaries(cell_range)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"Plage Excel invalide dans la configuration : {cell_range!r}"
        ) from exc
    if (
        any(isinstance(value, bool) or not isinstance(value, int) for value in bounds)
        or bounds[0] < 1
        or bounds[1] < 1
        or bounds[2] > EXCEL_MAX_COLUMN
        or bounds[3] > EXCEL_MAX_ROW
        or bounds[0] > bounds[2]
        or bounds[1] > bounds[3]
    ):
        raise ValueError(
            f"Plage Excel invalide dans la configuration : {cell_range!r}"
        )
    return bounds


def _tuple_of_strings(value: object, field_name: str) -> tuple[str, ...]:
    if value is None:
        return ()
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError(f"{field_name} doit être une liste de chaînes")
    return tuple(value)


def _validate_analysis(config: AnalysisConfig) -> None:
    if not config.extensions:
        raise ValueError("analysis.extensions doit contenir au moins une extension")
    unsupported_extensions = tuple(
        extension
        for extension in config.extensions
        if extension not in SUPPORTED_EXTENSIONS
    )
    if unsupported_extensions:
        raise ValueError(
            "Extension(s) Excel non prise(s) en charge dans analysis.extensions : "
            + ", ".join(unsupported_extensions)
        )
    for field_name in (
        "recursive",
        "case_sensitive_names",
        "strict_sheet_order",
        "use_stable_text_anchors",
        "detect_value_only_expansion",
        "report_ambiguities",
        "compare_formulas",
        "compare_dependencies",
        "compare_controlled_values",
    ):
        if not isinstance(getattr(config, field_name), bool):
            raise ValueError(f"analysis.{field_name} doit être un booléen TOML")
    for field_name in (
        "min_axis_similarity",
        "move_min_similarity",
        "ambiguity_margin",
        "numeric_absolute_tolerance",
        "numeric_relative_tolerance",
    ):
        value = getattr(config, field_name)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError(f"analysis.{field_name} doit être un nombre")
    for field_name in (
        "max_cells_per_sheet",
        "max_rows",
        "max_columns",
        "max_style_gap",
        "max_dependency_depth",
        "max_dependency_samples",
    ):
        value = getattr(config, field_name)
        if isinstance(value, bool) or not isinstance(value, int):
            raise ValueError(f"analysis.{field_name} doit être un entier")
    if not 0 <= config.min_axis_similarity <= 1:
        raise ValueError("analysis.min_axis_similarity doit être compris entre 0 et 1")
    if not 0 <= config.move_min_similarity <= 1:
        raise ValueError("analysis.move_min_similarity doit être compris entre 0 et 1")
    if config.move_min_similarity < config.min_axis_similarity:
        raise ValueError("analysis.move_min_similarity doit être supérieur ou égal à min_axis_similarity")
    if not 0 <= config.ambiguity_margin <= 1:
        raise ValueError("analysis.ambiguity_margin doit être compris entre 0 et 1")
    if config.max_cells_per_sheet < 1 or config.max_rows < 1 or config.max_columns < 1:
        raise ValueError("Les limites d'analyse doivent être strictement positives")
    if config.max_style_gap < 0:
        raise ValueError("analysis.max_style_gap doit être positif ou nul")
    if config.max_dependency_depth < 1 or config.max_dependency_samples < 1:
        raise ValueError("Les limites de dépendances doivent être strictement positives")
    if config.numeric_absolute_tolerance < 0 or config.numeric_relative_tolerance < 0:
        raise ValueError("Les tolérances numériques doivent être positives ou nulles")
    if not math.isfinite(config.numeric_absolute_tolerance) or not math.isfinite(
        config.numeric_relative_tolerance
    ):
        raise ValueError("Les tolérances numériques doivent être finies")


def load_config(path: str | Path | None = None) -> AppConfig:
    if path is None:
        return AppConfig()
    config_path = Path(path)
    if not config_path.is_file():
        raise FileNotFoundError(f"Fichier de configuration introuvable : {config_path}")
    with config_path.open("rb") as handle:
        raw = tomllib.load(handle)

    allowed_top = {"analysis", "sheet_rules"}
    unknown_top = set(raw) - allowed_top
    if unknown_top:
        raise ValueError(f"Clé(s) de configuration inconnue(s) : {', '.join(sorted(unknown_top))}")

    analysis_raw = raw.get("analysis", {})
    if not isinstance(analysis_raw, dict):
        raise ValueError("[analysis] doit être une table TOML")
    allowed_analysis = set(AnalysisConfig.__dataclass_fields__)
    unknown_analysis = set(analysis_raw) - allowed_analysis
    if unknown_analysis:
        raise ValueError(
            "Clé(s) [analysis] inconnue(s) : " + ", ".join(sorted(unknown_analysis))
        )
    if "extensions" in analysis_raw:
        extensions = _tuple_of_strings(analysis_raw["extensions"], "analysis.extensions")
        analysis_raw = dict(analysis_raw)
        normalized_extensions: list[str] = []
        for extension in extensions:
            normalized = extension.strip().lower()
            if not normalized:
                raise ValueError(
                    "analysis.extensions ne doit pas contenir d'extension vide"
                )
            if not normalized.startswith("."):
                normalized = f".{normalized}"
            if normalized not in normalized_extensions:
                normalized_extensions.append(normalized)
        analysis_raw["extensions"] = tuple(normalized_extensions)
    analysis = AnalysisConfig(**analysis_raw)
    _validate_analysis(analysis)

    rules_raw = raw.get("sheet_rules", [])
    if not isinstance(rules_raw, list):
        raise ValueError("[[sheet_rules]] doit être un tableau de tables TOML")
    rules: list[SheetRule] = []
    allowed_rule = set(SheetRule.__dataclass_fields__)
    for index, rule_raw in enumerate(rules_raw, start=1):
        if not isinstance(rule_raw, dict):
            raise ValueError(f"sheet_rules #{index} doit être une table TOML")
        unknown = set(rule_raw) - allowed_rule
        if unknown:
            raise ValueError(f"Clé(s) inconnue(s) dans sheet_rules #{index} : {', '.join(sorted(unknown))}")
        converted = dict(rule_raw)
        converted["editable_ranges"] = _tuple_of_strings(
            converted.get("editable_ranges"), f"sheet_rules #{index}.editable_ranges"
        )
        converted["formula_allowed_ranges"] = _tuple_of_strings(
            converted.get("formula_allowed_ranges"),
            f"sheet_rules #{index}.formula_allowed_ranges",
        )
        converted["monitored_ranges"] = _tuple_of_strings(
            converted.get("monitored_ranges"), f"sheet_rules #{index}.monitored_ranges"
        )
        converted["controlled_ranges"] = _tuple_of_strings(
            converted.get("controlled_ranges"), f"sheet_rules #{index}.controlled_ranges"
        )
        converted["critical_ranges"] = _tuple_of_strings(
            converted.get("critical_ranges"), f"sheet_rules #{index}.critical_ranges"
        )
        if not isinstance(converted.get("pattern", "*"), str):
            raise ValueError(f"sheet_rules #{index}.pattern doit être une chaîne")
        if "ignore" in converted and not isinstance(converted["ignore"], bool):
            raise ValueError(f"sheet_rules #{index}.ignore doit être un booléen TOML")
        rule = SheetRule(**converted)
        for configured_range in (
            rule.editable_ranges
            + rule.formula_allowed_ranges
            + rule.monitored_ranges
            + rule.controlled_ranges
            + rule.critical_ranges
        ):
            _contains(configured_range, 1, 1)
        for editable_range in rule.editable_ranges:
            for protected_range in rule.controlled_ranges + rule.critical_ranges:
                if _ranges_overlap(editable_range, protected_range):
                    raise ValueError(
                        f"sheet_rules #{index} contient des plages éditables et "
                        f"contrôlées qui se chevauchent : {editable_range!r} / "
                        f"{protected_range!r}"
                    )
        for formula_range in rule.formula_allowed_ranges:
            for protected_range in rule.controlled_ranges + rule.critical_ranges:
                if _ranges_overlap(formula_range, protected_range):
                    raise ValueError(
                        f"sheet_rules #{index} contient des formules autorisées et "
                        f"une zone contrôlée qui se chevauchent : {formula_range!r} / "
                        f"{protected_range!r}"
                    )
        rules.append(rule)
    return AppConfig(analysis=analysis, sheet_rules=tuple(rules))


def _ranges_overlap(first: str, second: str) -> bool:
    first_min_col, first_min_row, first_max_col, first_max_row = _range_bounds(first)
    second_min_col, second_min_row, second_max_col, second_max_row = _range_bounds(second)
    return not (
        first_max_row < second_min_row
        or second_max_row < first_min_row
        or first_max_col < second_min_col
        or second_max_col < first_min_col
    )
