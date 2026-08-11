from __future__ import annotations

import io
import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from html import escape
from html.parser import HTMLParser
from pathlib import Path

from popscheck.cli import main
from popscheck.reporting import generate_reports, render_country_html

from tests.helpers import create_workbook


class _CountryLinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.hrefs: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = dict(attrs)
        classes = set((attributes.get("class") or "").split())
        if tag == "a" and "country-link" in classes and attributes.get("href"):
            self.hrefs.append(attributes["href"] or "")


class ReportingAndCliP0Tests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def _run_cli(self, *arguments: str) -> tuple[int, str, str]:
        stdout, stderr = io.StringIO(), io.StringIO()
        with redirect_stdout(stdout), redirect_stderr(stderr):
            code = main(arguments)
        return code, stdout.getvalue(), stderr.getvalue()

    def test_category_counts_use_anomaly_code_before_incidental_message_words(self) -> None:
        rendered = render_country_html(
            {
                "country": "France",
                "reference_path": "sent/France.xlsx",
                "received_path": "received/France.xlsx",
                "anomalies": [
                    {
                        "code": "ROW_MOVED",
                        "message": "Ligne déplacée dans la feuille Données",
                        "impact": 2,
                    },
                    {
                        "code": "COLUMN_ADDED",
                        "message": "Colonne ajoutée dans la feuille Données",
                        "impact": 3,
                    },
                    {"code": "SHEET_REMOVED", "message": "Feuille retirée", "impact": 4},
                    {"code": "FILE_INVALID", "message": "Fichier invalide", "impact": 5},
                ],
            }
        )

        expected = {
            "feuilles": 4,
            "colonnes": 3,
            "lignes": 2,
            "fichier": 5,
            "autres": 0,
        }
        for category, count in expected.items():
            with self.subTest(category=category):
                self.assertIn(f'id="count-{category}">{count}</strong>', rendered)
        self.assertIn('id="total-anomalies">14</strong>', rendered)
        self.assertIn('data-anomaly-code="ROW_MOVED"', rendered)
        self.assertIn('data-category="lignes"', rendered)

    def test_html_escapes_all_result_supplied_text(self) -> None:
        payload = '<img src=x onerror="alert(1)"> & danger'
        rendered = render_country_html(
            {
                "country": payload,
                "reference_path": f"sent/{payload}.xlsx",
                "received_path": "received/ok.xlsx",
                "anomalies": [
                    {
                        "category": "lignes",
                        "code": "ROW_BAD",
                        "message": payload,
                        "details": {"<clé>": payload},
                    }
                ],
            },
            title=payload,
        )

        self.assertNotIn(payload, rendered)
        self.assertNotIn("<img src=x", rendered)
        self.assertIn(escape(payload, quote=True), rendered)
        self.assertIn("&lt;clé&gt;", rendered)

    def test_detail_report_shows_excel_location_sides_and_dimension_delta(self) -> None:
        rendered = render_country_html(
            {
                "country": "France",
                "reference_path": "sent/France.xlsx",
                "received_path": "received/France.xlsx",
                "status": "anomalies",
                "anomalies": [
                    {
                        "category": "colonnes",
                        "code": "COLUMN_REMOVED",
                        "message": "Colonne supprimée",
                        "sheet": "Forecast",
                        "location": "'Forecast'!D:D",
                        "expected_position": "D",
                        "expected": "COL::Mars",
                        "observed": "Absent",
                    }
                ],
                "metadata": {
                    "sheet_summaries": {
                        "Forecast": {
                            "expected_rows": 6,
                            "observed_rows": 6,
                            "row_delta": 0,
                            "expected_columns": 6,
                            "observed_columns": 5,
                            "column_delta": -1,
                        }
                    }
                },
            }
        )

        self.assertIn("Localisation Excel", rendered)
        self.assertIn(escape("'Forecast'!D:D", quote=True), rendered)
        self.assertIn("COL::Mars", rendered)
        self.assertIn("Absent", rendered)
        self.assertIn("Dimensions des feuilles affectées", rendered)
        self.assertIn('id="root-cause-count">1</strong>', rendered)
        self.assertIn('id="total-anomalies">1</strong>', rendered)

    def test_error_severity_promotes_the_overall_validation_badge(self) -> None:
        rendered = render_country_html(
            {
                "country": "France",
                "status": "anomalies",
                "anomalies": [
                    {
                        "category": "feuilles",
                        "code": "SHEET_REMOVED",
                        "message": "Feuille absente",
                        "severity": "error",
                    }
                ],
            }
        )

        self.assertIn(
            '<span class="status danger" id="status-global">Erreur structurelle</span>',
            rendered,
        )

    def test_generate_reports_makes_unique_slugs_and_only_real_detail_links(self) -> None:
        reports = self.root / "reports"
        countries = [
            {"key": "accent", "country": "Côte d'Ivoire"},
            {"key": "ascii", "country": "Cote d Ivoire"},
            {"key": "reserved", "country": "index"},
        ]

        index_path = generate_reports(countries, reports)

        self.assertEqual(reports / "index.html", index_path)
        self.assertTrue(index_path.is_file())
        details = {path.name for path in reports.glob("*.html")} - {"index.html"}
        self.assertEqual(3, len(details))
        self.assertIn("cote-d-ivoire.html", details)
        self.assertEqual(2, sum(name.startswith("cote-d-ivoire") for name in details))
        self.assertTrue(any(name.startswith("index-") for name in details))

        parser = _CountryLinkParser()
        parser.feed(index_path.read_text(encoding="utf-8"))
        self.assertEqual(details, set(parser.hrefs))
        for href in parser.hrefs:
            detail_path = reports / href
            self.assertTrue(detail_path.is_file(), href)
            self.assertIn('href="index.html"', detail_path.read_text(encoding="utf-8"))

    def test_cli_exit_codes_for_empty_missing_and_invalid_config(self) -> None:
        valid_config = self.root / "valid.toml"
        valid_config.write_text("", encoding="utf-8")

        empty_sent = self.root / "empty-sent"
        empty_received = self.root / "empty-received"
        empty_sent.mkdir()
        empty_received.mkdir()
        empty_reports = self.root / "empty-reports"
        code, _, stderr = self._run_cli(
            "--sent",
            str(empty_sent),
            "--received",
            str(empty_received),
            "--reports",
            str(empty_reports),
            "--config",
            str(valid_config),
        )
        self.assertEqual(0, code, stderr)
        self.assertTrue((empty_reports / "index.html").is_file())
        self.assertEqual([], json.loads((empty_reports / "resultats.json").read_text(encoding="utf-8"))["results"])

        missing_sent = self.root / "missing-sent"
        missing_received = self.root / "missing-received"
        missing_received.mkdir()
        create_workbook(missing_sent / "Missing.xlsx")
        code, _, stderr = self._run_cli(
            "--sent",
            str(missing_sent),
            "--received",
            str(missing_received),
            "--reports",
            str(self.root / "missing-reports"),
            "--config",
            str(valid_config),
            "--fail-on-issues",
        )
        self.assertEqual(1, code, stderr)

        invalid_config = self.root / "invalid.toml"
        invalid_config.write_text('[analysis]\nunknown_option = true\n', encoding="utf-8")
        code, _, stderr = self._run_cli(
            "--sent",
            str(empty_sent),
            "--received",
            str(empty_received),
            "--reports",
            str(self.root / "invalid-reports"),
            "--config",
            str(invalid_config),
        )
        self.assertEqual(2, code)
        self.assertIn("Erreur", stderr)


if __name__ == "__main__":
    unittest.main()
