from pathlib import Path
import sys
import tempfile
import unittest

import pandas as pd


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.product_candidate_sync import sync_shopify_catalog_enrichment
from modules.shopify_parser import (
    extract_variant_color_size,
    is_purchasing_color,
    normalize_order_export,
    parse_shopify_orders,
    sanitize_color_size,
)


class SizeColorGuardrailTests(unittest.TestCase):
    def test_size_and_fit_markers_are_never_variant_colors(self):
        cases = [
            (("Black", "14", "", "Black / 14"), ("Black", "14")),
            (("14", "TL", "", "14 / TL"), ("", "14")),
            (("3X TL", "", "", "3X TL"), ("", "3XLT")),
            (("Regular", "", "", "Regular"), ("", "")),
            (("Black", "3X TL", "", "Black / 3X TL"), ("Black", "3XLT")),
        ]
        for values, expected in cases:
            with self.subTest(values=values):
                self.assertEqual(extract_variant_color_size(*values), expected)

        for value in ("14", "Size 14", "14 Tall", "TL", "3X TL", "Regular", "11", "11-12.5"):
            with self.subTest(value=value):
                self.assertFalse(is_purchasing_color(value))
        self.assertTrue(is_purchasing_color("Burgundy"))

    def test_stale_size_in_color_is_recovered_when_it_is_a_real_size(self):
        self.assertEqual(sanitize_color_size("3X TL", ""), ("", "3XLT"))
        self.assertEqual(sanitize_color_size("TL", ""), ("", ""))
        self.assertEqual(sanitize_color_size("Black", "14"), ("Black", "14"))

    def test_shopify_import_keeps_a_true_color_and_assigns_numeric_apparel_size(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "numeric-apparel-size.csv"
            master_path = root / "product_master.csv"
            pd.DataFrame([{
                "Name": "#1901",
                "Lineitem name": "Women's Utility Pant - LP70",
                "Lineitem quantity": "1",
                "Lineitem option1": "Black",
                "Lineitem option2": "14",
                "Lineitem variant title": "Black / 14",
            }]).to_csv(csv_path, index=False)
            pd.DataFrame([{
                "Product Name": "Women's Utility Pant",
                "Style Number": "LP70",
                "Garment Color": "Black",
                "Vendor": "SanMar",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Product Category": "Pants / Jeans / Shorts",
                "Requires Size": "Yes",
                "Requires Color": "Yes",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            self.assertEqual(normalized.loc[0, "Variant Color"], "Black")
            self.assertEqual(normalized.loc[0, "Variant Size"], "14")
            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertEqual(parsed.loc[0, "Garment Color"], "Black")
            self.assertEqual(parsed.loc[0, "Size"], "14")

    def test_catalog_sync_refuses_a_false_size_color_from_a_legacy_event(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            master_path = Path(temporary_directory) / "product_master.csv"
            pd.DataFrame([{
                "Product Name": "Port Authority Polo",
                "Style Number": "K500",
                "Garment Color": "Black",
                "Vendor": "SanMar",
                "Decoration Type": "Embroidery",
                "Decoration Color": "Black",
                "Product Category": "Polos",
                "Requires Size": "Yes",
                "Requires Color": "Yes",
                "Requires Decoration": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            result = sync_shopify_catalog_enrichment(pd.DataFrame([{
                "Product Name": "Port Authority Polo",
                "Style Number": "K500",
                "Garment Color": "3X TL",
                "Parser Source": "Shopify Product",
            }]), master_path)
            master = pd.read_csv(master_path, dtype=str).fillna("")
            self.assertFalse(result["changed"])
            self.assertEqual(master["Garment Color"].tolist(), ["Black"])


if __name__ == "__main__":
    unittest.main()
