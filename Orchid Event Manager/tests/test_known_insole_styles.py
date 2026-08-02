from __future__ import annotations

import unittest

import pandas as pd

from modules.catalog_manager import clean_master
from modules.outsource_rules import is_boots_product, resolve_never_outsource
from modules.purchase_rules import infer_category, row_rules
from modules.review_workbook import _build_review_data
from modules.shopify_parser import normalize_order_export, parse_shopify_orders


INSOLE_STYLES = ("RKK0317", "A1Q82", "502440", "RKK0490")


class KnownInsoleStyleTests(unittest.TestCase):
    def test_exact_styles_are_boots_even_when_the_title_does_not_say_insole(self):
        for style in INSOLE_STYLES:
            with self.subTest(style=style):
                self.assertEqual(infer_category("Rocky Air-Port Footbed", style), "Boots")
                self.assertTrue(is_boots_product("Rocky Air-Port Footbed", "Other", style))

    def test_exact_styles_override_stale_color_and_decoration_requirements(self):
        for style in INSOLE_STYLES:
            with self.subTest(style=style):
                rules = row_rules({
                    "Product Name": "Rocky Air-Port Footbed",
                    "Style Number": style,
                    "Product Category": "Other",
                    "Requires Size": "Yes",
                    "Requires Color": "Yes",
                    "Requires Decoration": "Yes",
                    "Decoration Type": "Embroidery",
                })
                self.assertEqual(rules, {
                    "Product Category": "Boots",
                    "Requires Size": "Yes",
                    "Requires Color": "No",
                    "Requires Decoration": "No",
                })
                self.assertTrue(resolve_never_outsource("No", "Rocky Air-Port Footbed", "Other", style))

    def test_cleanup_repairs_legacy_product_master_records_permanently(self):
        master = pd.DataFrame([
            {
                "Product Name": "Rocky Air-Port Footbed",
                "Style Number": style,
                "Vendor": "Rocky Boots",
                "Garment Color": "",
                "Product Category": "Other",
                "Requires Size": "Yes",
                "Requires Color": "Yes",
                "Requires Decoration": "Yes",
                "Decoration Type": "Embroidery",
                "Decoration Location": "Left Chest",
                "Decoration Color": "Black",
                "Never Outsource": "No",
            }
            for style in INSOLE_STYLES
        ])

        cleaned = clean_master(master)
        self.assertEqual(set(cleaned["Product Category"]), {"Boots"})
        self.assertEqual(set(cleaned["Requires Size"]), {"Yes"})
        self.assertEqual(set(cleaned["Requires Color"]), {"No"})
        self.assertEqual(set(cleaned["Requires Decoration"]), {"No"})
        self.assertEqual(set(cleaned["Decoration Type"]), {"Blank Garment (No Decoration)"})
        self.assertEqual(set(cleaned["Decoration Location"]), {""})
        self.assertEqual(set(cleaned["Decoration Color"]), {""})
        self.assertEqual(set(cleaned["Never Outsource"]), {"Yes"})

    def test_purchase_review_does_not_request_garment_color_for_a_legacy_insole_record(self):
        """The screenshot case must clear without entering a fake garment color."""
        import tempfile
        from pathlib import Path

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "insole-sale.csv"
            master_path = root / "product_master.csv"
            pd.DataFrame([{
                "Order name": "#1734",
                "Created at": "2026-07-29 12:07:34",
                "Customer / Address / First name": "Kentavious",
                "Customer / Address / Last name": "Moss Line",
                "Billing address / Company": "Water",
                "Line item / Product title": "Rocky Air-Port Footbed - RKK0317",
                "Line item / Variant / Title": "2XLR 13-14",
                "Line item / Variant / Option 1": "2XLR 13-14",
                "Line item / Net quantity": "1",
                "Line item / Price": "20.00",
                "Line item / Discounts": "0",
                "Line item / Net sales": "20.00",
            }]).to_csv(csv_path, index=False)
            pd.DataFrame([{
                "Product Name": "Rocky Air-Port Footbed",
                "Style Number": "RKK0317",
                "Garment Color": "",
                "Vendor": "Rocky Boots",
                "Decoration Type": "Embroidery",
                "Decoration Location": "Left Chest",
                "Decoration Placement Instructions": "",
                "Decoration Color": "Black",
                "Product Category": "Other",
                "Requires Size": "Yes",
                "Requires Color": "Yes",
                "Requires Decoration": "Yes",
                "Never Outsource": "No",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertFalse(parsed.empty, parsed.to_dict(orient="records"))
            self.assertEqual(parsed.loc[0, "Style Number"], "RKK0317")
            self.assertEqual(parsed.loc[0, "Size"], "2XL")
            detail, review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )

            self.assertEqual(detail.loc[0, "Garment Color"], "")
            self.assertEqual(detail.loc[0, "Size"], "2XL")
            self.assertEqual(detail.loc[0, "Review Status"], "Ready")
            self.assertTrue(review.empty)


if __name__ == "__main__":
    unittest.main()
