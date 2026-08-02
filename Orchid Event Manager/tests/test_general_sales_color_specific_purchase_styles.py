"""Regression coverage for Orchid's normal vendor-PO workflow.

Never Outsource controls where goods ship.  It must not turn a normal General
Sales order into a Non-Included Items worksheet.  This module also proves the
new Product Master per-color purchasing-style mapping reaches the final PO.
"""

from __future__ import annotations

from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import pandas as pd


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules import final_po_generator
from modules.report_modes import GENERAL_SALES_PERIOD, UNIFORM_SIZING_EVENT
from modules.product_resolver import resolve_product


def boot_record(**overrides: object) -> dict[str, object]:
    record: dict[str, object] = {
        "Include": "Yes",
        "Purchase Vendor": "Timberland",
        "Decoration Type": "Blank Garment (No Decoration)",
        "Product #": "A64CQ",
        "Description": "Men's True Grit USA Pull-On Composite Toe Waterproof Work Boot",
        "Garment Color": "Black",
        "Size": "14 / Wide",
        "Quantity": 1,
        "Product Category": "Boots",
        "Requires Size": "Yes",
        "Requires Color": "No",
        "Requires Decoration": "No",
        "Do Not Outsource": "Yes",
        "Employee Name": "Chris Black",
        "Order Number": "#2001",
        "Source ID": "boot-black",
    }
    record.update(overrides)
    return record


class GeneralSalesAndColorSpecificPurchasingStyleTests(unittest.TestCase):
    def test_general_sales_never_outsource_boot_stays_on_its_vendor_po(self):
        with patch("modules.final_po_generator.load_review_lines", return_value=[boot_record()]):
            ready, review, non_included = final_po_generator._prepare_lines(
                Path("review.xlsx"), report_mode=GENERAL_SALES_PERIOD
            )

        self.assertTrue(review.empty)
        self.assertTrue(non_included.empty)
        self.assertEqual(ready["Vendor"].tolist(), ["Timberland"])
        self.assertEqual(ready["Garment Color"].tolist(), ["Black"])

    def test_uniform_sizing_event_keeps_the_condensed_ship_to_orchid_sheet(self):
        with patch("modules.final_po_generator.load_review_lines", return_value=[boot_record()]):
            ready, review, non_included = final_po_generator._prepare_lines(
                Path("review.xlsx"), report_mode=UNIFORM_SIZING_EVENT
            )

        self.assertTrue(review.empty)
        self.assertTrue(ready.empty)
        self.assertEqual(non_included["Vendor"].tolist(), ["Timberland"])

    def test_color_specific_purchase_style_flows_from_product_master_to_po(self):
        master = pd.DataFrame([
            {
                "Product Name": "24/7 Original Tactical Pant",
                "Style Number": "24/7",
                "Garment Color": "Black",
                "Color Aliases": "Black Full Grain",
                "Purchasing Style Number": "1062",
                "Vendor": "Tru-Spec",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Product Category": "Pants / Jeans / Shorts",
                "Requires Size": "Yes",
                "Requires Color": "Yes",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            },
            {
                "Product Name": "24/7 Original Tactical Pant",
                "Style Number": "24/7",
                "Garment Color": "Navy",
                "Color Aliases": "Navy Blue",
                "Purchasing Style Number": "1061",
                "Vendor": "Tru-Spec",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Product Category": "Pants / Jeans / Shorts",
                "Requires Size": "Yes",
                "Requires Color": "Yes",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            },
        ])

        black = resolve_product("24/7", "24/7 Original Tactical Pant", "Black Full Grain", master)
        navy = resolve_product("24/7", "24/7 Original Tactical Pant", "Navy Blue", master)
        self.assertEqual(black.garment_color, "Black")
        self.assertEqual(navy.garment_color, "Navy")
        self.assertEqual(black.as_dict()["Vendor Product #"], "1062")
        self.assertEqual(navy.as_dict()["Vendor Product #"], "1061")

        records = [
            boot_record(
                **{
                    "Purchase Vendor": "Tru-Spec",
                    "Product #": "24/7",
                    "Description": "24/7 Original Tactical Pant",
                    "Garment Color": black.garment_color,
                    "Vendor Product #": black.as_dict()["Vendor Product #"],
                    "Source ID": "truspec-black",
                }
            ),
            boot_record(
                **{
                    "Purchase Vendor": "Tru-Spec",
                    "Product #": "24/7",
                    "Description": "24/7 Original Tactical Pant",
                    "Garment Color": navy.garment_color,
                    "Vendor Product #": navy.as_dict()["Vendor Product #"],
                    "Source ID": "truspec-navy",
                }
            ),
        ]
        with patch("modules.final_po_generator.load_review_lines", return_value=records):
            ready, review, non_included = final_po_generator._prepare_lines(
                Path("review.xlsx"), report_mode=GENERAL_SALES_PERIOD
            )

        self.assertTrue(review.empty)
        self.assertTrue(non_included.empty)
        self.assertEqual(
            dict(zip(ready["Garment Color"], ready["Vendor Product #"])),
            {"Black": "1062", "Navy": "1061"},
        )


if __name__ == "__main__":
    unittest.main()
