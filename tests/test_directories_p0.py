from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from popscheck import analyze_directories
from popscheck.models import Status

from tests.helpers import create_workbook, file_fingerprint


class DirectoryAnalysisP0Tests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.sent = self.root / "sent"
        self.received = self.root / "received"
        self.sent.mkdir()
        self.received.mkdir()

    def test_missing_or_corrupt_file_does_not_abort_other_countries(self) -> None:
        create_workbook(self.sent / "France.xlsx")
        create_workbook(self.received / "France.xlsx", fill_inputs=True)
        create_workbook(self.sent / "Corrupt.xlsx")
        (self.received / "Corrupt.xlsx").write_bytes(b"not an Excel zip archive")
        create_workbook(self.sent / "Missing.xlsx")
        create_workbook(self.received / "Orphan.xlsx")

        results = {result.key: result for result in analyze_directories(self.sent, self.received)}

        self.assertEqual({"corrupt", "france", "missing", "orphan"}, set(results))
        self.assertEqual(Status.CONFORME, results["france"].status)
        self.assertEqual(Status.ERREUR, results["corrupt"].status)
        self.assertTrue(results["corrupt"].errors)
        self.assertEqual(Status.FICHIER_MANQUANT, results["missing"].status)
        self.assertEqual(Status.SANS_REFERENCE, results["orphan"].status)

    def test_normalized_name_collision_is_an_explicit_error(self) -> None:
        first = create_workbook(self.sent / "France.xlsx")
        shutil.copyfile(first, self.sent / "FRANCE.xlsm")
        create_workbook(self.received / "france.xlsx")

        results = analyze_directories(self.sent, self.received)

        self.assertEqual(1, len(results))
        self.assertEqual(Status.ERREUR, results[0].status)
        self.assertEqual("france", results[0].key)
        self.assertTrue(any("Plusieurs fichiers" in error for error in results[0].errors))

    def test_analysis_never_modifies_or_adds_files_in_input_directories(self) -> None:
        reference = create_workbook(self.sent / "France.xlsx")
        received = create_workbook(self.received / "France.xlsx", fill_inputs=True)
        before = {
            reference: file_fingerprint(reference),
            received: file_fingerprint(received),
        }
        directory_entries_before = {
            path.relative_to(self.root).as_posix()
            for path in self.root.rglob("*")
            if path.is_file()
        }

        results = analyze_directories(self.sent, self.received)

        self.assertEqual(Status.CONFORME, results[0].status)
        self.assertEqual(
            directory_entries_before,
            {
                path.relative_to(self.root).as_posix()
                for path in self.root.rglob("*")
                if path.is_file()
            },
        )
        self.assertEqual(before, {path: file_fingerprint(path) for path in before})


if __name__ == "__main__":
    unittest.main()
