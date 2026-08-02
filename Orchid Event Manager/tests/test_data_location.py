"""Regression checks for durable, non-iCloud Orchid data storage."""

from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from modules import paths


class DataLocationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.home = Path(self.temp.name) / "Rana"
        self.documents = paths.documents_data_dir(self.home)
        self.destination = paths.application_support_data_dir(self.home)

    def tearDown(self):
        self.temp.cleanup()

    def test_documents_data_is_copied_to_application_support_without_removal(self):
        report = self.documents / "reports" / "Completed Purchase Packets" / "Test 2" / "SanMar.pdf"
        state = self.documents / "current_review_state.json"
        report.parent.mkdir(parents=True)
        report.write_bytes(b"verified purchase order")
        state.write_text('{"event_name": "Test 2"}\n', encoding="utf-8")

        paths._migrate_documents_data(self.documents, self.destination)

        self.assertEqual(
            (self.destination / report.relative_to(self.documents)).read_bytes(),
            b"verified purchase order",
        )
        self.assertEqual(state.read_text(encoding="utf-8"), '{"event_name": "Test 2"}\n')
        marker = json.loads((self.destination / paths.DATA_LOCATION_MIGRATION_FILE).read_text(encoding="utf-8"))
        self.assertEqual(marker["source"], str(self.documents))

    def test_existing_application_support_data_is_never_overwritten(self):
        source_master = self.documents / "product_master.csv"
        target_master = self.destination / "product_master.csv"
        source_master.parent.mkdir(parents=True)
        self.destination.mkdir(parents=True)
        source_master.write_text("source data\n", encoding="utf-8")
        target_master.write_text("current local data\n", encoding="utf-8")

        paths._migrate_documents_data(self.documents, self.destination)

        self.assertEqual(target_master.read_text(encoding="utf-8"), "current local data\n")
        self.assertTrue((self.destination / paths.DATA_LOCATION_MIGRATION_FILE).is_file())

    def test_copied_active_event_state_uses_the_new_local_paths(self):
        state_path = self.documents / "current_review_state.json"
        report_path = self.documents / "reports" / "Purchase Orders" / "Test 2"
        state_path.parent.mkdir(parents=True)
        report_path.mkdir(parents=True)
        state_path.write_text(
            json.dumps(
                {
                    "last_purchase_order_dir": str(report_path),
                    "completed_purchase_order_archive_dir": str(self.documents / "reports" / "Event Archive" / "Test 2"),
                    "source_csv": "/Users/rana/Downloads/Test 2.csv",
                }
            ),
            encoding="utf-8",
        )

        paths._migrate_documents_data(self.documents, self.destination)

        copied = json.loads((self.destination / "current_review_state.json").read_text(encoding="utf-8"))
        self.assertEqual(
            copied["last_purchase_order_dir"],
            str(self.destination / "reports" / "Purchase Orders" / "Test 2"),
        )
        self.assertEqual(
            copied["completed_purchase_order_archive_dir"],
            str(self.destination / "reports" / "Event Archive" / "Test 2"),
        )
        self.assertEqual(copied["source_csv"], "/Users/rana/Downloads/Test 2.csv")

    def test_default_location_is_not_documents(self):
        self.assertEqual(
            paths.application_support_data_dir(self.home),
            self.home / "Library" / "Application Support" / paths.APP_DATA_FOLDER,
        )
        self.assertNotIn("Documents", str(paths.application_support_data_dir(self.home)))


if __name__ == "__main__":
    unittest.main()
