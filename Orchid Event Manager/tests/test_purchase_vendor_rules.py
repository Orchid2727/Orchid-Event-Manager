"""Regression checks for safe internal/non-included supplier routing."""

from __future__ import annotations

from pathlib import Path
import sys
import unittest


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.purchase_vendor_rules import is_internal_purchase_vendor
from modules import final_po_generator


class PurchaseVendorRulesTests(unittest.TestCase):
    def test_orchid_is_never_a_purchase_supplier(self):
        for value in ("Orchid", "Orchid Uniforms & Apparel", "orchid uniforms and apparel"):
            with self.subTest(value=value):
                self.assertTrue(is_internal_purchase_vendor(value))

    def test_actual_suppliers_remain_valid_purchase_vendors(self):
        for value in ("Georgia Boot", "Rocky Boots", "SanMar", "VF"):
            with self.subTest(value=value):
                self.assertFalse(is_internal_purchase_vendor(value))

    def test_internal_orchid_vendor_is_returned_to_review_not_printed_as_a_supplier(self):
        record = {
            "Include": "Yes", "Purchase Vendor": "Orchid Uniforms & Apparel",
            "Decoration Type": "Blank Garment (No Decoration)", "Product #": "A437Y",
            "Description": "Men's True Grit Work Boot", "Garment Color": "",
            "Size": "11 / Medium", "Quantity": 1, "Product Category": "Boots",
            "Requires Size": "Yes", "Requires Color": "No", "Requires Decoration": "No",
            "Do Not Outsource": "Yes", "Employee Name": "Alex Hernandez", "Order Number": "#1714",
        }
        original_loader = final_po_generator.load_review_lines
        final_po_generator.load_review_lines = lambda _path: [record]
        try:
            ready, review, non_included = final_po_generator._prepare_lines(Path("review.xlsx"))
        finally:
            final_po_generator.load_review_lines = original_loader

        self.assertTrue(ready.empty)
        self.assertTrue(non_included.empty)
        self.assertEqual(len(review), 1)
        self.assertIn("select the actual supplier", review.iloc[0]["Review Reason"])


if __name__ == "__main__":
    unittest.main()
