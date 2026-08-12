from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from popscheck.config import AppConfig, SheetRule, load_config


class ConfigSemanticsTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def test_non_finite_numeric_tolerance_is_rejected(self) -> None:
        config_path = self.root / "popscheck.toml"
        config_path.write_text(
            "[analysis]\nnumeric_absolute_tolerance = inf\n",
            encoding="utf-8",
        )

        with self.assertRaisesRegex(ValueError, "doivent être finies"):
            load_config(config_path)

    def test_protected_ranges_override_cross_rule_permissions(self) -> None:
        config = AppConfig(
            sheet_rules=(
                SheetRule(
                    pattern="*",
                    editable_ranges=("B2:B2",),
                    formula_allowed_ranges=("B2:B2",),
                ),
                SheetRule(pattern="Budget", critical_ranges=("B2:B2",)),
            )
        )

        rule = config.rule_for("Budget")

        self.assertFalse(rule.cell_is_editable(2, 2))
        self.assertFalse(rule.formula_is_allowed(2, 2))
        self.assertTrue(rule.cell_is_monitored(2, 2))

    def test_invalid_or_unbounded_excel_ranges_are_rejected_cleanly(self) -> None:
        for configured_range in ("B2:A1", "A0:B2", "XFE1:XFE2", "A:A"):
            with self.subTest(configured_range=configured_range):
                config_path = self.root / "popscheck.toml"
                config_path.write_text(
                    "[[sheet_rules]]\n"
                    'pattern = "Budget"\n'
                    f'controlled_ranges = ["{configured_range}"]\n',
                    encoding="utf-8",
                )

                with self.assertRaisesRegex(ValueError, "Plage Excel invalide"):
                    load_config(config_path)

    def test_extensions_are_normalized_deduplicated_and_validated(self) -> None:
        config_path = self.root / "popscheck.toml"
        config_path.write_text(
            '[analysis]\nextensions = [" XLSX ", ".xlsx", "xlsm"]\n',
            encoding="utf-8",
        )

        config = load_config(config_path)

        self.assertEqual((".xlsx", ".xlsm"), config.analysis.extensions)

        for invalid in ('[]', '[""]', '["csv"]'):
            with self.subTest(invalid=invalid):
                config_path.write_text(
                    f"[analysis]\nextensions = {invalid}\n",
                    encoding="utf-8",
                )
                with self.assertRaises(ValueError):
                    load_config(config_path)


if __name__ == "__main__":
    unittest.main()
