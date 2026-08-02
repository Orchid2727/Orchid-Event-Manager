"""Regression coverage for orphaned Purchase Review recovery entries."""

from __future__ import annotations

import importlib
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))


class DecorationAuditOrphanRecoveryTests(unittest.TestCase):
    def _load_app(self):
        """Import the non-UI application logic without requiring desktop Tk."""
        saved_ctk = sys.modules.get("customtkinter")
        saved_app = sys.modules.pop("app", None)
        sys.modules["customtkinter"] = types.SimpleNamespace(CTk=type("CTk", (), {}))

        def restore_modules():
            sys.modules.pop("app", None)
            if saved_app is not None:
                sys.modules["app"] = saved_app
            if saved_ctk is None:
                sys.modules.pop("customtkinter", None)
            else:
                sys.modules["customtkinter"] = saved_ctk

        self.addCleanup(restore_modules)
        return importlib.import_module("app")

    def _write_workbook(self, path: Path) -> None:
        headers = [
            "Line ID", "Source ID", "Decision Key", "Include", "Product #", "Description",
            "Garment Color", "Size", "Quantity", "Purchase Vendor", "Order Number",
            "Employee Name", "Decoration Type", "Decoration Location", "Decoration Color",
            "Review Status", "Review Reason", "Action Required", "Fix In",
        ]
        values = [
            "line-current", "source-current", "decision-current", "Yes", "PC55T", "Port & Company Tee",
            "Navy", "L", 2, "SanMar", "#1001", "Avery Employee", "Embroidery",
            "Left Chest", "", "Ready", "", "", "",
        ]
        workbook = Workbook()
        review = workbook.active
        review.title = "Review & Edit"
        review.append(headers)
        review.append(values)
        source = workbook.create_sheet("All PO Lines")
        source.append(headers)
        source.append(values)
        workbook.save(path)
        workbook.close()

    def test_orphaned_saved_review_decision_cannot_blank_the_color_audit(self):
        """A deleted older line must not hide a still-current embroidery color row."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            previous_data_dir = os.environ.get("ORCHID_DATA_DIR")
            os.environ["ORCHID_DATA_DIR"] = str(root / "orchid-data")
            self.addCleanup(
                lambda: os.environ.__setitem__("ORCHID_DATA_DIR", previous_data_dir)
                if previous_data_dir is not None
                else os.environ.pop("ORCHID_DATA_DIR", None)
            )
            app = self._load_app()

            workbook_path = root / "purchase_review.xlsx"
            self._write_workbook(workbook_path)
            master_path = root / "product_master.csv"
            pd.DataFrame([{
                "Product Name": "Port & Company Tee",
                "Style Number": "PC55T",
                "Garment Color": "Navy",
                "Vendor": "SanMar",
                "Product Category": "T-Shirt",
                "Decoration Type": "Embroidery",
                "Decoration Location": "Left Chest",
                "Decoration Color": "White",
            }]).to_csv(master_path, index=False)

            manager = object.__new__(app.OrchidPurchaseManager)
            manager.last_review_workbook = workbook_path
            manager._decoration_audit_event_records_cache_key = None
            manager._decoration_audit_event_records_cache = []
            manager._completed_review_records_from_state = lambda _workbook: [{"job_id": "old"}]
            manager._audit_override_jobs_for_workbook = lambda _workbook: []
            manager._journal_completion_jobs_for_workbook = lambda _workbook: [
                {
                    "job_id": "deleted-line",
                    "workbook": str(workbook_path),
                    "source_id": "source-deleted",
                    "line_id": "line-deleted",
                    "decision_key": "decision-deleted",
                    "values": {
                        "Product #": "OLD999",
                        "Garment Color": "Black",
                        "Size": "M",
                        "Quantity": "1",
                        "Purchase Vendor": "SanMar",
                    },
                    "issue": {"order": "#999", "employee": "Former Employee"},
                    "allow_missing": False,
                    "require_event_match": True,
                },
                {
                    "job_id": "current-line",
                    "workbook": str(workbook_path),
                    "source_id": "source-current",
                    "line_id": "line-current",
                    "decision_key": "decision-current",
                    "values": {
                        "Decoration Color": "White",
                        "Decoration Location": "Left Chest",
                    },
                    "issue": {"order": "#1001", "employee": "Avery Employee"},
                    "allow_missing": False,
                    "require_event_match": True,
                },
            ]

            records = manager._decoration_audit_records_from_current_event(force=True)
            audit = app.build_decoration_color_audit(records, master_path, {})

            self.assertEqual(len(audit), 1)
            self.assertEqual(audit[0]["style"], "PC55T")
            self.assertEqual(audit[0]["current_color"], "White")

            reconciled = root / "orchid-data" / "reconciled_review_workbooks" / (
                workbook_path.stem + "_journal_reconciled.xlsx"
            )
            saved = load_workbook(reconciled)
            try:
                for sheet_name in ("Review & Edit", "All PO Lines"):
                    sheet = saved[sheet_name]
                    headers = {cell.value: column for column, cell in enumerate(sheet[1], start=1)}
                    self.assertEqual(sheet.cell(2, headers["Decoration Color"]).value, "White")
            finally:
                saved.close()

    def test_missing_source_id_is_orphaned_even_when_old_key_matches_a_new_line(self):
        """A stale immutable source ID cannot be revived by a broad old key.

        This is the exact failure that produced a phantom ``1 to Review`` with
        zero audit rows: the old key was present on a different regenerated
        line, but Source ID matching intentionally refused to apply the saved
        decision.  The job is therefore historical and must not block the
        Decoration Color Audit.
        """
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            previous_data_dir = os.environ.get("ORCHID_DATA_DIR")
            os.environ["ORCHID_DATA_DIR"] = str(root / "orchid-data")
            self.addCleanup(
                lambda: os.environ.__setitem__("ORCHID_DATA_DIR", previous_data_dir)
                if previous_data_dir is not None
                else os.environ.pop("ORCHID_DATA_DIR", None)
            )
            app = self._load_app()

            workbook_path = root / "purchase_review.xlsx"
            self._write_workbook(workbook_path)
            job = {
                "job_id": "unsafe-to-skip",
                "workbook": str(workbook_path),
                "source_id": "source-no-longer-matches",
                # This existing line is evidence that the job may still apply.
                "line_id": "line-current",
                "decision_key": "decision-current",
                "values": {"Decoration Color": "White"},
                "issue": {"order": "#1001", "employee": "Avery Employee"},
                "allow_missing": False,
                "require_event_match": True,
            }

            skipped = app.OrchidPurchaseManager._apply_review_save_jobs(
                [job], skip_orphaned_recovery_jobs=True
            )
            self.assertEqual(
                skipped,
                [{
                    "job_id": "unsafe-to-skip",
                    "source_id": "source-no-longer-matches",
                    "line_id": "line-current",
                    "decision_key": "decision-current",
                }],
            )

            saved = load_workbook(workbook_path)
            try:
                for sheet_name in ("Review & Edit", "All PO Lines"):
                    sheet = saved[sheet_name]
                    headers = {cell.value: column for column, cell in enumerate(sheet[1], start=1)}
                    self.assertIsNone(sheet.cell(2, headers["Decoration Color"]).value)
            finally:
                saved.close()


if __name__ == "__main__":
    unittest.main()
