from __future__ import annotations

import unittest
from collections import Counter

from popscheck.align import AxisItem, align_axis, axis_similarity
from popscheck.config import AnalysisConfig


class AxisSimilarityTests(unittest.TestCase):
    def _item(self, formula: str, *, label_is_unique: bool) -> AxisItem:
        label = "L|Total"
        tokens = Counter(
            {
                label: 1,
                f"F|{formula}": 1,
                f"P|{formula}": 1,
            }
        )
        return AxisItem(
            index=1,
            key=formula,
            tokens=tokens,
            label="Total",
            unique_labels=frozenset({label}) if label_is_unique else frozenset(),
        )

    def test_unique_label_keeps_axis_matched_when_formula_logic_changes(self) -> None:
        expected = self._item("=A1+B1", label_is_unique=True)
        observed = self._item("=A1-B1", label_is_unique=True)

        self.assertEqual(1.0, axis_similarity(expected, observed))

    def test_repeated_label_does_not_mask_formula_identity_change(self) -> None:
        expected = self._item("=A1+B1", label_is_unique=False)
        observed = self._item("=A1-B1", label_is_unique=False)

        self.assertLess(axis_similarity(expected, observed), 0.62)

    def test_report_ambiguities_false_suppresses_all_alignment_warnings(self) -> None:
        items = [
            AxisItem(
                index=index,
                key="same",
                tokens=Counter({"L|Repeated": 1}),
                information=10.0,
            )
            for index in range(1, 551)
        ]

        alignment = align_axis(
            items,
            items,
            AnalysisConfig(report_ambiguities=False),
        )

        self.assertEqual([], alignment.ambiguities)


if __name__ == "__main__":
    unittest.main()
