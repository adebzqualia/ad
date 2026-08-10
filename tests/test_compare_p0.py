from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from popscheck import compare_workbooks
from popscheck.models import Status

from tests.helpers import (
    BASE_COLUMNS,
    BASE_ROWS,
    BASE_SHEETS,
    create_workbook,
    insert_after,
    move_after,
    without,
)


class StructuralComparisonP0Tests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def _compare(
        self,
        *,
        reference_options: dict[str, object] | None = None,
        **received_options: object,
    ):
        reference = create_workbook(
            self.root / "sent" / "France.xlsx", **(reference_options or {})
        )
        received = create_workbook(
            self.root / "received" / "France.xlsx", **received_options
        )
        return compare_workbooks(reference, received)

    def _compare_axis(self, **received_options: object):
        one_sheet = ("Données",)
        return self._compare(
            reference_options={"sheets": one_sheet},
            sheets=one_sheet,
            **received_options,
        )

    def assert_only_code(self, result, expected_code: str) -> None:
        self.assertEqual(Status.ANOMALIES, result.status, result.to_dict())
        self.assertEqual([expected_code], [item.code for item in result.anomalies])

    def test_entered_values_do_not_create_structural_anomalies(self) -> None:
        result = self._compare(fill_inputs=True)

        self.assertEqual(Status.CONFORME, result.status, result.to_dict())
        self.assertEqual([], result.anomalies)
        self.assertEqual(0, result.total_anomalies)

    def test_sheet_added_removed_and_reordered(self) -> None:
        cases = (
            (
                "added",
                (*BASE_SHEETS, "Notes pays"),
                "SHEET_ADDED",
            ),
            (
                "removed",
                without(BASE_SHEETS, "Prévisions"),
                "SHEET_REMOVED",
            ),
            (
                "reordered",
                ("Synthèse", "Prévisions", "Données"),
                "SHEET_ORDER_CHANGED",
            ),
        )
        for label, sheet_order, expected_code in cases:
            with self.subTest(label=label):
                result = self._compare(sheets=sheet_order)
                self.assert_only_code(result, expected_code)

    def test_added_row_is_reported_at_its_observed_position(self) -> None:
        result = self._compare_axis(
            rows=insert_after(BASE_ROWS, "Coûts", "Ajustement pays")
        )

        self.assert_only_code(result, "ROW_ADDED")
        anomaly = result.anomalies[0]
        self.assertEqual("lignes", anomaly.category)
        self.assertEqual(4, anomaly.observed_position)
        self.assertIsNone(anomaly.expected_position)

    def test_removed_row_is_reported_at_its_expected_position(self) -> None:
        result = self._compare_axis(rows=without(BASE_ROWS, "Marge"))

        self.assert_only_code(result, "ROW_REMOVED")
        anomaly = result.anomalies[0]
        self.assertEqual("lignes", anomaly.category)
        self.assertEqual(4, anomaly.expected_position)
        self.assertIsNone(anomaly.observed_position)

    def test_moved_row_reports_expected_and_observed_positions(self) -> None:
        result = self._compare_axis(
            rows=move_after(BASE_ROWS, "Coûts", "Prévision")
        )

        self.assert_only_code(result, "ROW_MOVED")
        anomaly = result.anomalies[0]
        self.assertEqual("lignes", anomaly.category)
        self.assertEqual(3, anomaly.expected_position)
        self.assertEqual(5, anomaly.observed_position)

    def test_added_column_is_reported_at_its_observed_position(self) -> None:
        result = self._compare_axis(
            columns=insert_after(BASE_COLUMNS, "Février", "Ajustement pays")
        )

        self.assert_only_code(result, "COLUMN_ADDED")
        anomaly = result.anomalies[0]
        self.assertEqual("colonnes", anomaly.category)
        self.assertEqual("D", anomaly.observed_position)
        self.assertIsNone(anomaly.expected_position)

    def test_removed_column_is_reported_at_its_expected_position(self) -> None:
        result = self._compare_axis(columns=without(BASE_COLUMNS, "Mars"))

        self.assert_only_code(result, "COLUMN_REMOVED")
        anomaly = result.anomalies[0]
        self.assertEqual("colonnes", anomaly.category)
        self.assertEqual("D", anomaly.expected_position)
        self.assertIsNone(anomaly.observed_position)

    def test_moved_column_reports_expected_and_observed_positions(self) -> None:
        result = self._compare_axis(
            columns=move_after(BASE_COLUMNS, "Février", "Avril")
        )

        self.assert_only_code(result, "COLUMN_MOVED")
        anomaly = result.anomalies[0]
        self.assertEqual("colonnes", anomaly.category)
        self.assertEqual("C", anomaly.expected_position)
        self.assertEqual("E", anomaly.observed_position)


if __name__ == "__main__":
    unittest.main()
