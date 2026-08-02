"""Regression checks for final purchase-order PDF discovery."""

from __future__ import annotations

from pathlib import Path
import unittest

from modules.purchase_order_files import is_final_purchase_order_pdf


class PurchaseOrderFileClassificationTests(unittest.TestCase):
    def test_ship_to_orchid_pdf_is_a_final_purchase_order(self):
        self.assertTrue(
            is_final_purchase_order_pdf("Ship_to_Orchid_Purchase_Order.pdf")
        )

    def test_legacy_non_included_filename_stays_available(self):
        self.assertTrue(
            is_final_purchase_order_pdf("Non-Included_Items_Purchase_Order.pdf")
        )

    def test_vendor_purchase_order_is_a_final_purchase_order(self):
        self.assertTrue(is_final_purchase_order_pdf("Timberland_Purchase_Order.pdf"))

    def test_supporting_reports_are_not_final_purchase_orders(self):
        for name in (
            "Screen_Print_Job.pdf",
            "Employee_Totals.pdf",
            "Event_Summary.pdf",
            "Internal_Purchase_Order.pdf",
        ):
            with self.subTest(name=name):
                self.assertFalse(is_final_purchase_order_pdf(name))

    def test_generation_opens_the_verified_folder_after_confirmation(self):
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        source = app_path.read_text(encoding="utf-8")
        generation_start = source.index("    def _run_final_purchase_order_generation")
        generation_end = source.index("    def create_final_purchase_orders", generation_start)
        generation = source[generation_start:generation_end]
        self.assertIn("self.after(120, self.open_latest_purchase_orders)", generation)


if __name__ == "__main__":
    unittest.main()
