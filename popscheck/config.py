from __future__ import annotations

import fnmatch
import tomllib
from dataclasses import dataclass, field, replace
from pathlib import Path

from openpyxl.utils.cell import range_boundaries


SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")


@dataclass(frozen=True, slots=True)
class SheetRule:
    pattern: str = "*"
    ignore: bool = False
    editable_ranges: tuple[str, ...] = ()
    monitored_ranges: tuple[str, ...] = ()

    def matches(self, sheet_name: str) -> bool:
        return fnmatch.fnmatchcase(sheet_name.casefold(), self.pattern.casefold())

    def cell_is_editable(self, row: int, column: int) -> bool:
        return any(_contains(cell_range, row, column) for cell_range in self.editable_ranges)

    def cell_is_monitored(self, row: int, column: int) -> bool:
        if not self.monitored_ranges:
            return True
        return any(_contains(cell_range, row, column) for cell_range in self.monitored_ranges)

    def axis_coordinate_is_monitored(self, axis: str, index: int) -> bool:
        """Check one known axis while the other coordinate may still be unknown."""

        if not self.monitored_ranges:
            return True
        for configured_range in self.monitored_ranges:
            min_col, min_row, max_col, max_row = range_boundaries(configured_range)
            if axis == "row" and min_row <= index <= max_row:
                return True
            if axis == "column" and min_col <= index <= max_col:
                return True
        return False

    def monitored_extent(self) -> tuple[int, int] | None:
        if not self.monitored_ranges:
            return None
        max_row = 0
        max_column = 0
        for configured_range in self.monitored_ranges:
            _min_col, _min_row, range_max_col, range_max_row = range_boundaries(configured_range)
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
                    monitored_ranges=selected.monitored_ranges + rule.monitored_ranges,
                )
        return selected


def _contains(cell_range: str, row: int, column: int) -> bool:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Plage Excel invalide dans la configuration : {cell_range!r}") from exc
    return min_row <= row <= max_row and min_col <= column <= max_col


def _tuple_of_strings(value: object, field_name: str) -> tuple[str, ...]:
    if value is None:
        return ()
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError(f"{field_name} doit être une liste de chaînes")
    return tuple(value)


def _validate_analysis(config: AnalysisConfig) -> None:
    for field_name in (
        "recursive",
        "case_sensitive_names",
        "strict_sheet_order",
        "use_stable_text_anchors",
        "detect_value_only_expansion",
        "report_ambiguities",
    ):
        if not isinstance(getattr(config, field_name), bool):
            raise ValueError(f"analysis.{field_name} doit être un booléen TOML")
    for field_name in (
        "min_axis_similarity",
        "move_min_similarity",
        "ambiguity_margin",
    ):
        value = getattr(config, field_name)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError(f"analysis.{field_name} doit être un nombre")
    for field_name in (
        "max_cells_per_sheet",
        "max_rows",
        "max_columns",
        "max_style_gap",
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
        analysis_raw["extensions"] = tuple(
            extension.lower() if extension.startswith(".") else f".{extension.lower()}"
            for extension in extensions
        )
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
        converted["monitored_ranges"] = _tuple_of_strings(
            converted.get("monitored_ranges"), f"sheet_rules #{index}.monitored_ranges"
        )
        if not isinstance(converted.get("pattern", "*"), str):
            raise ValueError(f"sheet_rules #{index}.pattern doit être une chaîne")
        if "ignore" in converted and not isinstance(converted["ignore"], bool):
            raise ValueError(f"sheet_rules #{index}.ignore doit être un booléen TOML")
        rule = SheetRule(**converted)
        for configured_range in rule.editable_ranges + rule.monitored_ranges:
            _contains(configured_range, 1, 1)
        rules.append(rule)
    return AppConfig(analysis=analysis, sheet_rules=tuple(rules))
