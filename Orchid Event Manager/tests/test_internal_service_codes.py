"""Regression checks for Orchid-only service and decoration charge codes."""

from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest

import pandas as pd

APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.internal_services import (
    is_decoration_charge,
    is_in_house_service_product,
    is_internal_service_style,
    is_phantom_product_label,
)
from modules.purchase_rules import row_rules
from modules.review_workbook import _build_review_data
from modules.shopify_parser import normalize_order_export, parse_shopify_orders
from modules.shopify_parser import is_decoration_service
from modules.catalog_manager import clean_master


class InternalServiceCodeTests(unittest.TestCase):
    def test_confirmed_orchid_service_codes_are_never_treated_as_garments(self):
        # These are Orchid's live, internal add-on/decoration charge products.
        # They must not request a color or size and must not create a vendor PO.
        for style in (
            "750",
            "750B",
            "BAGDECORATION",
            "EMBARK",
            "EMBARK2LOGOS",
            "FLAG",
            "AMERICANFLAG",
            "PARKSLEAFLOGOONLY",
            "800",
        ):
            with self.subTest(style=style):
                self.assertTrue(is_internal_service_style(style))
                self.assertTrue(is_decoration_service("Embark", style))
                self.assertEqual(
                    row_rules({
                        "Style Number": style,
                        "Product Name": "Embark",
                        "Product Category": "Other",
                    }),
                    {
                        "Product Category": "Other",
                        "Requires Size": "No",
                        "Requires Color": "No",
                        "Requires Decoration": "No",
                    },
                )

    def test_normal_logo_garment_style_remains_purchaseable(self):
        self.assertFalse(is_internal_service_style("K500"))
        self.assertFalse(is_decoration_service("Port Authority Silk Touch Polo", "K500"))

    def test_higher_stitch_count_750b_charge_never_enters_product_master_or_review(self):
        """750B is a price add-on, not an item Orchid needs to buy or set up."""
        product_name = "Higher Stitch Count Embroidered Left Chest Logo"
        self.assertTrue(is_internal_service_style("750B"))
        self.assertTrue(is_decoration_charge(product_name))
        self.assertTrue(is_decoration_service(product_name, "750B"))

        self.assertTrue(
            is_in_house_service_product(product_name, style_number="750B")
        )

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "higher-stitch-charge.csv"
            master_path = root / "product_master.csv"
            pd.DataFrame([{
                "Order name": "#1800",
                "Created at": "2026-07-30 12:00:00",
                "Line item / Product title": product_name,
                "Line item / SKU": "750B",
                "Line item / Net quantity": "1",
                "Line item / Price": "10.00",
                "Line item / Discounts": "0",
                "Line item / Net sales": "10.00",
            }]).to_csv(csv_path, index=False)
            pd.DataFrame().to_csv(master_path, index=False)
            self.assertTrue(parse_shopify_orders(csv_path, master_path).empty)

    def test_embark_two_logos_service_line_never_enters_purchase_review(self):
        """The #1747 screenshot must clear without entering a fake color or size."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "embark-service.csv"
            master_path = root / "product_master.csv"
            pd.DataFrame([{
                "Order name": "#1747",
                "Created at": "2026-07-29 12:07:34",
                "Customer / Address / First name": "James",
                "Customer / Address / Last name": "Smith",
                "Billing address / Company": "Embark",
                "Line item / Product title": "Embark 2 Logos",
                "Line item / SKU": "EMBARK2LOGOS",
                "Line item / Net quantity": "2",
                "Line item / Price": "0.00",
                "Line item / Discounts": "0",
                "Line item / Net sales": "0.00",
            }]).to_csv(csv_path, index=False)
            pd.DataFrame(columns=[
                "Product Name", "Style Number", "Garment Color", "Vendor",
                "Decoration Type", "Decoration Location", "Decoration Color",
                "Product Category", "Requires Size", "Requires Color",
                "Requires Decoration", "Never Outsource",
            ]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertTrue(parsed.empty, parsed.to_dict(orient="records"))
            detail, review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            self.assertTrue(detail.empty)
            self.assertTrue(review.empty)

    def test_parks_leaf_logo_charge_without_sku_never_enters_product_master(self):
        """Parks Leaf Logo Only is a decoration fee, not a purchasable item."""
        for product_name in ("Parks Leaf Logo", "Parks Leaf Logo Only"):
            with self.subTest(product_name=product_name):
                self.assertTrue(is_decoration_charge(product_name))
                self.assertTrue(is_decoration_service(product_name))
                self.assertTrue(is_in_house_service_product(product_name))

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "parks-leaf-logo-charge.csv"
            master_path = root / "product_master.csv"
            pd.DataFrame([{
                "Order name": "#1900",
                "Created at": "2026-07-31 12:00:00",
                "Line item / Product title": "Parks Leaf Logo Only",
                "Line item / Net quantity": "1",
                "Line item / Price": "10.00",
                "Line item / Discounts": "0",
                "Line item / Net sales": "10.00",
            }]).to_csv(csv_path, index=False)
            pd.DataFrame().to_csv(master_path, index=False)
            self.assertTrue(parse_shopify_orders(csv_path, master_path).empty)

    def test_bag_decoration_charge_without_sku_never_enters_or_remains_in_product_master(self):
        """The checkout charge can arrive as the bare title ``Bag Decoration``."""
        product_name = "Bag Decoration"
        self.assertTrue(is_decoration_charge(product_name))
        self.assertTrue(is_decoration_service(product_name))
        self.assertTrue(is_in_house_service_product(product_name))

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "bag-decoration-charge.csv"
            master_path = root / "product_master.csv"
            pd.DataFrame([{
                "Order name": "#1901",
                "Created at": "2026-07-31 12:00:00",
                "Line item / Product title": product_name,
                "Line item / Net quantity": "1",
                "Line item / Price": "10.00",
                "Line item / Discounts": "0",
                "Line item / Net sales": "10.00",
            }]).to_csv(csv_path, index=False)
            pd.DataFrame().to_csv(master_path, index=False)
            self.assertTrue(parse_shopify_orders(csv_path, master_path).empty)

        stale_master = pd.DataFrame([{
            "Product Name": product_name,
            "Style Number": "",
            "Garment Color": "",
            "Vendor": "",
            "Decoration Type": "Embroidery",
        }])
        self.assertTrue(clean_master(stale_master).empty)

    def test_new_service_and_bulk_stock_items_never_become_product_master_candidates(self):
        """Current event rows must not keep recreating excluded Orchid products."""
        cases = (
            ("EMBARK", "Embark", "Embark logo decoration charge"),
            ("AMERICANFLAG", "American Flag", "American Flag"),
            ("EMBARK2LOGOS", "Embark 2 Logo Embroidery", "Embark 2 Logo Embroidery"),
            ("800", "Belt", "Bulk stock belt"),
        )
        for style, title, original_line in cases:
            with self.subTest(style=style):
                self.assertTrue(is_internal_service_style(style))
                self.assertTrue(is_decoration_service(original_line, style))
                self.assertTrue(is_in_house_service_product(title, original_line, style_number=style))

    def test_medium_black_is_never_a_product_master_item(self):
        self.assertTrue(is_phantom_product_label("Medium Black"))
        self.assertTrue(is_phantom_product_label("MEDIUMBLACK"))
        self.assertFalse(is_phantom_product_label("Black Ascent Pant"))
        self.assertTrue(is_decoration_service("Medium Black"))


if __name__ == "__main__":
    unittest.main()
