"""Regression checks for the internal receiving worksheet layout."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

import pandas as pd
from pypdf import PdfReader

from modules.final_po_generator import (
    _build_in_house_receiving_report,
    _vendor_po_number,
)
from modules.mission_control import _live_routes


class InHouseReceivingReportTests(unittest.TestCase):
    def test_saved_vendor_po_replaces_manual_routing_label(self):
        po_numbers = {
            ("berne apparel", "combined vendor order"): "Berne Line 2026",
        }
        self.assertEqual(
            _vendor_po_number(po_numbers, "Berne Apparel"),
            "Berne Line 2026",
        )

    def test_manual_receiving_vendor_appears_in_po_assignment(self):
        routes = _live_routes(
            Path("/nonexistent.xlsx"),
            "Uniform Sizing Event",
            set(),
            None,
            lines=[{
                "Include": "Yes",
                "Do Not Outsource": "Yes",
                "Purchase Vendor": "Berne Apparel",
                "Decoration Type": "Blank Garments",
                "Quantity": 2,
                "Line ID": "berne-1",
            }],
        )
        self.assertEqual(len(routes), 1)
        self.assertEqual(routes[0]["vendor"], "Berne Apparel")
        self.assertEqual(routes[0]["report_type"], "Consolidated Vendor Order")

    def test_report_uses_full_width_vendor_heading_without_total_footer(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "receiving.pdf"
            detail = pd.DataFrame([{
                "Vendor": "Berne Apparel",
                "Assigned PO Number": "Berne Line 2026",
                "Style Number": "B415",
                "Product Name": "Men's Heritage Insulated Bib Overall",
                "Garment Color": "Black",
                "Size": "M",
                "Quantity": 1,
                "Employee Name": "Alex Hernandez",
                "Order Number": "#1714",
                "Operational Decoration Type": "Blank Garment",
                "Operational Decoration Location": "",
                "Operational Decoration Color": "",
                "Notes": "",
            }])
            _build_in_house_receiving_report(
                detail, output, "Test 2", "Standard Orchid Workflow"
            )
            text = PdfReader(str(output)).pages[0].extract_text()

        self.assertIn("Berne Apparel", text)
        self.assertIn("PO # Berne Line 2026", text)
        self.assertNotIn("VENDOR TOTAL", text)
        self.assertIn("Test 2", text)
        self.assertIn("RECEIVING & DECORATION REPORT", text)
        self.assertNotIn("IN-HOUSE RECEIVING & DECORATION REPORT", text)
        self.assertNotIn("MANUAL / NON-INCLUDED", text)
        self.assertNotIn("1 PIECES", text)
        self.assertLess(text.index("Received"), text.index("Product #"))

    def test_receiving_report_has_full_width_rows_and_no_vendor_total_bar(self):
        import inspect
        from modules import final_po_generator

        source = inspect.getsource(_build_in_house_receiving_report)
        self.assertIn("report_width = page_width - left_margin - right_margin", source)
        self.assertIn("width=report_width", source)
        self.assertIn("3.01*inch", source)
        self.assertNotIn("VENDOR TOTAL", source)

    def test_receiving_report_draws_its_masthead_only_on_the_first_page(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "receiving-multi-page.pdf"
            detail = pd.DataFrame([{
                "Vendor": "Berne Apparel",
                "Assigned PO Number": "Berne Line 2026",
                "Style Number": f"B{index:03}",
                "Product Name": "Men's Heritage Insulated Bib Overall",
                "Garment Color": "Black",
                "Size": "M",
                "Quantity": 1,
                "Employee Name": f"Employee {index}",
                "Order Number": f"#{1700 + index}",
                "Operational Decoration Type": "Blank Garment",
                "Operational Decoration Location": "",
                "Operational Decoration Color": "",
                "Notes": "",
            } for index in range(90)])
            _build_in_house_receiving_report(
                detail, output, "Test 2", "Standard Orchid Workflow"
            )
            reader = PdfReader(str(output))
            page_text = [page.extract_text() or "" for page in reader.pages]

        self.assertGreater(len(page_text), 1)
        self.assertIn("RECEIVING & DECORATION REPORT", page_text[0])
        self.assertNotIn("RECEIVING & DECORATION REPORT", page_text[1])

    def test_receiving_report_uses_the_shared_po_masthead_on_its_first_page(self):
        import inspect
        from modules import final_po_generator

        source = inspect.getsource(_build_in_house_receiving_report)
        self.assertIn("_draw_internal_document_header", source)
        self.assertIn("autoNextPageTemplate=\"receiving-later-pages\"", source)
        self.assertNotIn("onLaterPages=draw_header", source)
        self.assertIn("RECEIVING & DECORATION REPORT", source)
        self.assertIn("letter_center_from_bottom", inspect.getsource(final_po_generator._draw_internal_document_header))

    def test_receiving_report_uses_vendor_product_number_when_available(self):
        """The receiving sheet must match the Tru-Spec number on the PO."""
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "receiving-truspec.pdf"
            detail = pd.DataFrame([{
                "Vendor": "Tru-Spec",
                "Assigned PO Number": "TRU-TEST-1",
                "Style Number": "1555",
                "Vendor Product #": "1556",
                "Product Name": "Pro Vector Pant",
                "Garment Color": "Coyote",
                "Size": "32x30",
                "Quantity": 1,
                "Employee Name": "Alex Hernandez",
                "Order Number": "#1714",
                "Operational Decoration Type": "Blank Garment",
                "Operational Decoration Location": "",
                "Operational Decoration Color": "",
                "Notes": "",
            }])
            _build_in_house_receiving_report(
                detail, output, "Test 2", "Standard Orchid Workflow"
            )
            text = PdfReader(str(output)).pages[0].extract_text()

        self.assertIn("1556", text)
        self.assertNotIn("1555", text)


if __name__ == "__main__":
    unittest.main()
