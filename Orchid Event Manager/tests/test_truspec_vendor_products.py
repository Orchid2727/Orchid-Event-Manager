from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import pandas as pd
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle

APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.final_po_generator import _compact_table, _prepare_lines
from modules.truspec_vendor_products import default_vendor_product_number


class TruSpecVendorProductTests(unittest.TestCase):
    def test_screenshot_families_use_the_confirmed_color_specific_numbers(self):
        cases = [
            ("Pro Vector Pant", "Black", "1560"),
            ("Pro Vector Pant", "Coyote", "1556"),
            ("24/7 Original Tactical Pant Rip-Stop", "Black", "1062"),
            ("24/7 Original Tactical Pant Rip-Stop", "Brown", "1065"),
            ("24/7 Original Tactical Pant Rip-Stop", "Coyote", "1063"),
            ("24/7 Original Tactical Pant Rip-Stop", "Grey", "1089"),
            ("24/7 Original Tactical Pant Rip-Stop", "Khaki", "1060"),
            ("24/7 Original Tactical Pant Rip-Stop", "Navy", "1061"),
            ("24/7 Original Tactical Pant Rip-Stop", "Olive Drab", "1064"),
            ("Agility Pant", "Black", "1526"),
            ("Agility Pant", "Flat Dark Earth", "1528"),
        ]
        for product, color, expected in cases:
            with self.subTest(product=product, color=color):
                self.assertEqual(
                    default_vendor_product_number("Tru-Spec", product, color), expected
                )

    def test_stale_known_color_number_is_corrected_but_custom_override_survives(self):
        self.assertEqual(
            default_vendor_product_number("Tru-Spec", "Pro Vector Pant", "Coyote", "1555"),
            "1556",
        )
        self.assertEqual(
            default_vendor_product_number("Tru-Spec", "Agility Pant", "Black", "CUSTOM-1526"),
            "CUSTOM-1526",
        )

    def test_vendor_po_displays_vendor_number_instead_of_shared_orchid_style(self):
        body = ParagraphStyle("TruSpecBody")
        quantity = ParagraphStyle("TruSpecQty", parent=body, alignment=TA_CENTER)
        header = ParagraphStyle("TruSpecHeader", parent=body)
        detail = pd.DataFrame([{
            "Vendor Product #": "1062", "Style Number": "24/7",
            "Product Name": "24/7 Original Tactical Pant Rip-Stop", "Garment Color": "Black",
            "Size": "34x32", "Company": "Orchid", "Employee Name": "Test",
            "Order Number": "#1", "Quantity": 2,
        }])
        table = _compact_table(detail, body, quantity, header)
        self.assertEqual(table._cellvalues[1][0].getPlainText(), "1062")

    def test_final_report_pipeline_replaces_generic_review_styles(self):
        records = [
            {
                "Include": "Yes", "Purchase Vendor": "Tru-Spec",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Product Category": "Pants / Jeans / Shorts",
                "Requires Size": "Yes", "Requires Color": "Yes", "Requires Decoration": "No",
                "Product #": "1555", "Description": "Pro Vector Pant",
                "Garment Color": "Coyote", "Size": "32x30", "Quantity": "1",
                "Source ID": "vector-coyote",
            },
            {
                "Include": "Yes", "Purchase Vendor": "Tru-Spec",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Product Category": "Pants / Jeans / Shorts",
                "Requires Size": "Yes", "Requires Color": "Yes", "Requires Decoration": "No",
                "Product #": "24/7", "Description": "24/7 Original Tactical Pant Rip-Stop",
                "Garment Color": "Olive Drab", "Size": "34x32", "Quantity": "1",
                "Source ID": "original-olive",
            },
            {
                "Include": "Yes", "Purchase Vendor": "Tru-Spec",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Product Category": "Pants / Jeans / Shorts",
                "Requires Size": "Yes", "Requires Color": "Yes", "Requires Decoration": "No",
                "Product #": "TRU-SPEC", "Description": "Agility Pant",
                "Garment Color": "Flat Dark Earth", "Size": "32x30", "Quantity": "1",
                "Source ID": "agility-fde",
            },
        ]
        with patch("modules.final_po_generator.load_review_lines", return_value=records):
            ready, review, non_included = _prepare_lines(Path("review.xlsx"))
        self.assertTrue(review.empty)
        self.assertTrue(non_included.empty)
        numbers = dict(zip(ready["Source ID"], ready["Vendor Product #"]))
        self.assertEqual(numbers, {
            "vector-coyote": "1556",
            "original-olive": "1064",
            "agility-fde": "1528",
        })


if __name__ == "__main__":
    unittest.main()
