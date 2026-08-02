import sys
import tempfile
import unittest
import inspect
from pathlib import Path

import pandas as pd
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch

APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.final_po_generator import (
    REPORT_HEADER_RESERVE,
    REPORT_HEADER_RULE_FROM_TOP,
    REPORT_JOB_NAME_BAND_HEIGHT,
    _build_pdf,
    _compact_table,
)


class VendorPurchaseOrderLayoutTests(unittest.TestCase):
    def test_quantity_header_fits_on_one_line(self):
        body_style = ParagraphStyle("Body", fontName="Helvetica", fontSize=9.4, leading=10.6)
        qty_style = ParagraphStyle("Qty", parent=body_style, alignment=TA_CENTER)
        header_style = ParagraphStyle(
            "Header", parent=body_style, fontName="Helvetica-Bold", fontSize=9.3,
            leading=10.4, textColor=colors.white, alignment=TA_LEFT,
        )
        detail = pd.DataFrame([{
            "Style Number": "MQK00023",
            "Product Name": "Clique Ice Pique Mens Short Sleeve Tech Polo",
            "Garment Color": "Silver",
            "Size": "L",
            "Company": "Fleet Services",
            "Employee Name": "Jerry Johnson",
            "Order Number": "#1700",
            "Quantity": 1,
        }])

        table = _compact_table(detail, body_style, qty_style, header_style)
        quantity_column_width = table._colWidths[4]
        self.assertGreaterEqual(quantity_column_width, 0.70 * inch)

        quantity_header = table._cellvalues[0][4]
        _width, height = quantity_header.wrap(quantity_column_width - 10, 100)
        self.assertLessEqual(height, header_style.leading)

    def test_vendor_led_header_has_no_legacy_po_number_box(self):
        """Vendor POs retain the approved one-header layout on every rebuild."""
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "vendor_purchase_order.pdf"
            detail = pd.DataFrame([{
                "Decoration Type": "Embroidery",
                "Decoration Location": "Left Chest",
                "Decoration Color": "Black",
                "Style Number": "MEBRAN",
                "Product Name": "Men's Epic Easy Care 3/4-Sleeve Top",
                "Garment Color": "Black",
                "Size": "M",
                "Company": "Utilities",
                "Employee Name": "Alex Hernandez",
                "Order Number": "#1700",
                "Quantity": 4,
            }])
            _build_pdf(
                detail, output, "PO-123", "Cutter & Buck", "Uniform Sizing Event",
                event_name="Test 2",
            )
            reader = PdfReader(str(output))
            text = reader.pages[0].extract_text()

        self.assertIn("PURCHASE ORDER", text)
        self.assertIn("Cutter & Buck", text)
        self.assertIn("Test 2", text)
        self.assertIn("EMBROIDERY - LEFT CHEST", text)
        self.assertNotIn("PO NUMBER", text)

    def test_vendor_title_uses_the_orchid_letter_centerline(self):
        """Keep the vendor title and wordmark letters on one optical line."""
        source = inspect.getsource(_build_pdf)
        self.assertIn("letter_center_from_bottom", source)
        self.assertIn("wordmark_y = wordmark_letter_center_y - letter_center_from_bottom", source)
        self.assertIn("vendor_center_y = wordmark_letter_center_y", source)

    def test_job_name_band_is_one_third_shorter_on_all_shared_purchase_headers(self):
        """Do not let any purchase worksheet quietly regain the tall job band."""
        former_band_height = 0.92 * inch
        self.assertAlmostEqual(REPORT_JOB_NAME_BAND_HEIGHT, former_band_height * (2.0 / 3.0))
        self.assertAlmostEqual(
            REPORT_HEADER_RESERVE,
            REPORT_HEADER_RULE_FROM_TOP + REPORT_JOB_NAME_BAND_HEIGHT,
        )


if __name__ == "__main__":
    unittest.main()
