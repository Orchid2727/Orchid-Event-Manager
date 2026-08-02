"""Regression checks for the audited, test-only Purchase Review bypass."""

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules import final_po_generator


def _unresolved_record() -> dict:
    return {
        "Source ID": "test-source-1",
        "Line ID": "test-line-1",
        "Original Quantity": 2,
        "Original Order Number": "#1001",
        "Original Company": "Test Department",
        "Original Employee Name": "Taylor Test",
        "Quantity Override Confirmed": "No",
        "Include": "Yes",
        "Purchase Vendor": "SanMar",
        "Product #": "K500",
        "Description": "Port Authority Silk Touch Polo",
        "Original Shopify Line": "Port Authority Silk Touch Polo",
        "Garment Color": "Black",
        "Size": "M",
        "Quantity": 2,
        "Company": "Test Department",
        "Employee Name": "Taylor Test",
        "Order Number": "#1001",
        "Requires Size": "Yes",
        "Requires Color": "Yes",
        "Requires Decoration": "No",
        "Decoration Type": "Blank Garment (No Decoration)",
        "Decoration Location": "",
        "Decoration Color": "",
        "Do Not Outsource": "No",
        "Review Status": "Needs Review",
        "Event Review Reason": "Customer decision required",
    }


class TestPurchaseOrderBypassTests(unittest.TestCase):
    def test_test_mode_holds_unresolved_review_lines_out_of_vendor_orders(self):
        record = _unresolved_record()
        with patch.object(final_po_generator, "load_review_lines", return_value=[record]):
            ready, held, non_included = final_po_generator._prepare_lines(
                Path("review.xlsx"), hold_unresolved_review=True,
            )
        self.assertTrue(ready.empty)
        self.assertTrue(non_included.empty)
        self.assertEqual(len(held), 1)
        self.assertEqual(held.loc[0, "Review Reason"], "Customer decision required")

    def test_normal_generation_keeps_existing_behavior_and_hold_report_is_readable(self):
        record = _unresolved_record()
        with patch.object(final_po_generator, "load_review_lines", return_value=[record]):
            ready, held, non_included = final_po_generator._prepare_lines(
                Path("review.xlsx"), hold_unresolved_review=False,
            )
        self.assertEqual(len(ready), 1)
        self.assertTrue(held.empty)
        self.assertTrue(non_included.empty)

        with tempfile.TemporaryDirectory() as temporary_directory:
            pdf_path = Path(temporary_directory) / "TEST_ONLY_Held_for_Review.pdf"
            with patch.object(final_po_generator, "load_review_lines", return_value=[record]):
                _ready, held, _non_included = final_po_generator._prepare_lines(
                    Path("review.xlsx"), hold_unresolved_review=True,
                )
            final_po_generator._build_test_review_hold_pdf(held, pdf_path, "Large CSV Test")
            self.assertTrue(pdf_path.is_file())
            rendered = "\n".join(page.extract_text() or "" for page in PdfReader(str(pdf_path)).pages)
            self.assertIn("HELD FROM PURCHASE ORDERS", rendered)
            self.assertIn("Taylor Test", rendered)
            self.assertIn("Customer decision required", rendered)


if __name__ == "__main__":
    unittest.main()
