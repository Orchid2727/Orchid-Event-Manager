"""Regression checks for the non-included-items report header and total."""

from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest

import pandas as pd
from pypdf import PdfReader

APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.final_po_generator import _build_non_included_items_pdf


class NonIncludedReportLayoutTests(unittest.TestCase):
    def test_job_name_vendor_po_and_vendor_total_are_compact(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "non_included.pdf"
            detail = pd.DataFrame([{
                "Vendor": "Berne Apparel",
                "Style Number": "B415",
                "Product Name": "Men's Heritage Insulated Bib Overall",
                "Garment Color": "Black",
                "Size": "M",
                "Company": "Wastewater",
                "Employee Name": "Alex Hernandez",
                "Order Number": "#1714",
                "Quantity": 2,
            }])
            _build_non_included_items_pdf(
                detail,
                output,
                "Test 2",
                {("berne apparel", "combined vendor order"): "Berne Line 2026"},
            )
            text = PdfReader(str(output)).pages[0].extract_text()

        self.assertIn("Test 2", text)
        self.assertIn("Berne Apparel", text)
        self.assertIn("PO # Berne Line 2026", text)
        self.assertIn("VENDOR TOTAL: 2", text)
        self.assertIn("SHIP TO ORCHID PURCHASE ORDER", text)
        self.assertNotIn("INTERNAL PURCHASE ORDER", text)
        self.assertNotIn("ORDER MANUALLY AND SHIP TO ORCHID", text)

    def test_assigned_po_number_is_preferred_for_vendor_heading(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "non_included.pdf"
            detail = pd.DataFrame([{
                "Vendor": "Berne Apparel",
                "Assigned PO Number": "Berne Line 2026",
                "Style Number": "B415",
                "Product Name": "Men's Heritage Insulated Bib Overall",
                "Garment Color": "Black",
                "Size": "M",
                "Company": "Wastewater",
                "Employee Name": "Alex Hernandez",
                "Order Number": "#1714",
                "Quantity": 2,
            }])
            _build_non_included_items_pdf(
                detail,
                output,
                "Test 2",
                {("berne apparel", "combined vendor order"): "Older PO"},
            )
            text = PdfReader(str(output)).pages[0].extract_text()

        self.assertIn("PO # Berne Line 2026", text)
        self.assertNotIn("PO # Older PO", text)

    def test_identical_purchase_lines_combine_while_employee_orders_remain_visible(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "non_included.pdf"
            detail = pd.DataFrame([
                {
                    "Vendor": "Georgia Boot",
                    "Vendor Product #": "GB00111",
                    "Style Number": "GB00111",
                    "Product Name": "Comfort Insole",
                    "Garment Color": "",
                    "Size": "L",
                    "Company": "Utilities",
                    "Employee Name": "Alex Hernandez",
                    "Order Number": "#1714",
                    "Quantity": 1,
                },
                {
                    "Vendor": "Georgia Boot",
                    "Vendor Product #": "GB00111",
                    "Style Number": "GB00111",
                    "Product Name": "Comfort Insole",
                    "Garment Color": "",
                    "Size": "L",
                    "Company": "Utilities",
                    "Employee Name": "Blair Taylor",
                    "Order Number": "#1715",
                    "Quantity": 1,
                },
            ])
            _build_non_included_items_pdf(detail, output, "Test 2")
            text = PdfReader(str(output)).pages[0].extract_text()

        self.assertEqual(text.count("GB00111"), 1)
        self.assertIn("2", text)
        self.assertIn("Alex Hernandez", text)
        self.assertIn("Blair Taylor", text)
        self.assertIn("#1714", text)
        self.assertIn("#1715", text)

    def test_boot_width_is_visible_and_kept_separate_when_aggregating(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "non_included_boots.pdf"
            detail = pd.DataFrame([
                {
                    "Vendor": "Georgia Boot", "Vendor Product #": "GB00752",
                    "Style Number": "GB00752", "Product Name": "Aero LT Work Boot",
                    "Garment Color": "", "Size": "11 / Medium", "Company": "Water",
                    "Employee Name": "Alex Hernandez", "Order Number": "#1714", "Quantity": 1,
                },
                {
                    "Vendor": "Georgia Boot", "Vendor Product #": "GB00752",
                    "Style Number": "GB00752", "Product Name": "Aero LT Work Boot",
                    "Garment Color": "", "Size": "11 / Wide", "Company": "Water",
                    "Employee Name": "Blair Taylor", "Order Number": "#1715", "Quantity": 1,
                },
                {
                    "Vendor": "Georgia Boot", "Vendor Product #": "GB00752",
                    "Style Number": "GB00752", "Product Name": "Aero LT Work Boot",
                    "Garment Color": "", "Size": "11 / Medium", "Company": "Water",
                    "Employee Name": "Casey Morgan", "Order Number": "#1716", "Quantity": 1,
                },
            ])
            _build_non_included_items_pdf(detail, output, "Line Maintenance Boots 2026")
            text = PdfReader(str(output)).pages[0].extract_text()

        self.assertIn("Size / Width", text)
        self.assertIn("11 / Medium", text)
        self.assertIn("11 / Wide", text)
        # Two Medium boots combine; the Wide boot stays separate.
        self.assertIn("2", text)
        self.assertIn("3", text)

    def test_a64cq_color_alias_results_stay_separate_in_the_pdf(self):
        """A boot with an optional color still needs its selected color printed."""
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "non_included_a64cq_colors.pdf"
            detail = pd.DataFrame([
                {
                    "Vendor": "Timberland", "Style Number": "A64CQ",
                    "Product Name": "True Grit USA Pull-On Work Boot",
                    "Garment Color": "Black", "Size": "14 / Wide", "Company": "Water",
                    "Employee Name": "Chris Black", "Order Number": "#2001", "Quantity": 2,
                },
                {
                    "Vendor": "Timberland", "Style Number": "A64CQ",
                    "Product Name": "True Grit USA Pull-On Work Boot",
                    "Garment Color": "Medium Brown", "Size": "14 / Wide", "Company": "Water",
                    "Employee Name": "Pat Brown", "Order Number": "#2002", "Quantity": 3,
                },
            ])
            _build_non_included_items_pdf(detail, output, "Line Boots 2026")
            text = PdfReader(str(output)).pages[0].extract_text()

        self.assertIn("Black", text)
        self.assertIn("Medium Brown", text)
        self.assertIn("Chris Black", text)
        self.assertIn("Pat Brown", text)

    def test_non_included_report_uses_the_shared_po_masthead_on_later_pages(self):
        import inspect
        from modules import final_po_generator

        source = inspect.getsource(_build_non_included_items_pdf)
        self.assertIn("_draw_internal_document_header", source)
        self.assertIn("onLaterPages=draw_header", source)
        self.assertIn("SHIP TO ORCHID PURCHASE ORDER", source)
        self.assertIn("letter_center_from_bottom", inspect.getsource(final_po_generator._draw_internal_document_header))


if __name__ == "__main__":
    unittest.main()
