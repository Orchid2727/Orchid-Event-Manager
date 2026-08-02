from pathlib import Path
import sys
import tempfile
import unittest

import pandas as pd


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.review_workbook import _apply_previous_review_edits, _build_review_data, _revalidate_detail
from modules.shopify_parser import (
    extract_variant_color_size,
    normalize_order_export,
    parse_shopify_orders,
)
from modules.decoration_color_audit import build_decoration_color_audit


class InsoleSizeVariantTests(unittest.TestCase):
    def test_shoe_fit_range_keeps_only_the_selected_insole_size(self):
        cases = {
            "Medium 6-8 1/2": "M",
            "Large 9-10 1/2": "L",
            "XL 11-12 1/2": "XL",
            "2XL 13-14 1/2": "2XL",
            "2XLR 13-14": "2XL",
            "ME 6–8.5": "M",
            "LR 9–10.5": "L",
        }
        for option, expected_size in cases.items():
            with self.subTest(option=option):
                self.assertEqual(
                    extract_variant_color_size(option, "", "", option),
                    ("", expected_size),
                )

    def test_boot_size_and_width_stay_together_for_purchasing(self):
        self.assertEqual(
            extract_variant_color_size("11", "Regular", "", "11 / Regular"),
            ("", "11 / Medium"),
        )
        self.assertEqual(
            extract_variant_color_size("10.5", "Wide", "", "10.5 / Wide"),
            ("", "10.5 / Wide"),
        )

    def test_optional_boot_colors_are_preserved_and_split_by_color_alias(self):
        """A color optional boot still prints Shopify-selected purchasing colors."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "a64cq-boot-sales.csv"
            master_path = root / "product_master.csv"

            pd.DataFrame([
                {
                    "Order name": "#2001",
                    "Created at": "2026-07-30 08:00:00",
                    "Customer / Address / First name": "Chris",
                    "Customer / Address / Last name": "Black",
                    "Billing address / Company": "Water",
                    "Line item / Product title": "Men's True Grit USA Pull-On Composite Toe Waterproof Work Boot - A64CQ",
                    "Line item / Variant / Title": "14 / Wide / Black Full Grain",
                    "Line item / Variant / Option 1": "14",
                    "Line item / Variant / Option 2": "Wide",
                    "Line item / Variant / Option 3": "Black Full Grain",
                    "Line item / Variant / SKU": "A64CQ",
                    "Line item / Net quantity": "2",
                    "Line item / Price": "252.94",
                    "Line item / Discounts": "0",
                    "Line item / Net sales": "505.88",
                },
                {
                    "Order name": "#2002",
                    "Created at": "2026-07-30 08:01:00",
                    "Customer / Address / First name": "Pat",
                    "Customer / Address / Last name": "Brown",
                    "Billing address / Company": "Water",
                    "Line item / Product title": "Men's True Grit USA Pull-On Composite Toe Waterproof Work Boot - A64CQ",
                    "Line item / Variant / Title": "12 / Wide / Medium Brown Full Grain",
                    "Line item / Variant / Option 1": "12",
                    "Line item / Variant / Option 2": "Wide",
                    "Line item / Variant / Option 3": "Medium Brown Full Grain",
                    "Line item / Variant / SKU": "A64CQ",
                    "Line item / Net quantity": "3",
                    "Line item / Price": "252.94",
                    "Line item / Discounts": "0",
                    "Line item / Net sales": "758.82",
                },
            ]).to_csv(csv_path, index=False)
            pd.DataFrame([
                {
                    "Product Name": "Men's True Grit USA Pull-On Composite Toe Waterproof Work Boot",
                    "Style Number": "A64CQ",
                    "Garment Color": "Black",
                    "Color Aliases": "Black Full Grain",
                    "Vendor": "Timberland",
                    "Decoration Type": "Blank Garment (No Decoration)",
                    "Decoration Location": "",
                    "Decoration Color": "",
                    "Product Category": "Boots",
                    "Requires Size": "Yes",
                    "Requires Color": "No",
                    "Requires Decoration": "No",
                    "Never Outsource": "Yes",
                    "Setup Required": "No",
                },
                {
                    "Product Name": "Men's True Grit USA Pull-On Composite Toe Waterproof Work Boot",
                    "Style Number": "A64CQ",
                    "Garment Color": "Medium Brown",
                    "Color Aliases": "Medium Brown Full Grain",
                    "Vendor": "Timberland",
                    "Decoration Type": "Blank Garment (No Decoration)",
                    "Decoration Location": "",
                    "Decoration Color": "",
                    "Product Category": "Boots",
                    "Requires Size": "Yes",
                    "Requires Color": "No",
                    "Requires Decoration": "No",
                    "Never Outsource": "Yes",
                    "Setup Required": "No",
                },
            ]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            self.assertEqual(normalized["Variant Color"].tolist(), ["Black Full Grain", "Medium Brown Full Grain"])
            self.assertEqual(normalized["Variant Size"].tolist(), ["14 / Wide", "12 / Wide"])

            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertEqual(parsed["Garment Color"].tolist(), ["Black Full Grain", "Medium Brown Full Grain"])

            detail, review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            self.assertEqual(detail["Garment Color"].tolist(), ["Black", "Medium Brown"])
            self.assertEqual(detail["Size"].tolist(), ["14 / Wide", "12 / Wide"])
            self.assertTrue(review.empty)

            refreshed = _revalidate_detail(detail)
            self.assertEqual(refreshed["Garment Color"].tolist(), ["Black", "Medium Brown"])
            self.assertEqual(refreshed["Review Status"].tolist(), ["Ready", "Ready"])

    def test_report_toaster_insole_selection_is_ready_without_a_false_color(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "boot-sales.csv"
            master_path = root / "product_master.csv"

            pd.DataFrame([{
                "Order name": "#1766",
                "Created at": "2026-07-29 12:07:34",
                "Customer / Address / First name": "Chris",
                "Customer / Address / Last name": "Williams",
                "Billing address / Company": "Water",
                "Line item / Product title": "Georgia Boot CC5 Insole - GB00111",
                "Line item / Variant / Title": "XL 11-12 1/2",
                "Line item / Variant / Option 1": "XL 11-12 1/2",
                "Line item / Net quantity": "2",
                "Line item / Price": "29.41",
                "Line item / Discounts": "0",
                "Line item / Net sales": "58.82",
            }]).to_csv(csv_path, index=False)

            pd.DataFrame([{
                "Product Name": "Georgia Boot CC5 Insole",
                "Style Number": "GB00111",
                "Garment Color": "",
                "Vendor": "Georgia Boot",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Decoration Location": "",
                "Decoration Placement Instructions": "",
                "Decoration Color": "",
                "Product Category": "Boots",
                "Requires Size": "Yes",
                "Requires Color": "No",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            self.assertEqual(normalized.loc[0, "Variant Color"], "")
            self.assertEqual(normalized.loc[0, "Variant Size"], "XL")

            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertEqual(parsed.loc[0, "Style Number"], "GB00111")
            self.assertEqual(parsed.loc[0, "Garment Color"], "")
            self.assertEqual(parsed.loc[0, "Size"], "XL")

            detail, review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            self.assertEqual(detail.loc[0, "Size"], "XL")
            self.assertEqual(detail.loc[0, "Garment Color"], "")
            self.assertEqual(detail.loc[0, "Review Status"], "Ready")
            self.assertTrue(review.empty)

    def test_abbreviated_decimal_insole_variant_is_ready_without_a_false_color(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "boot-sales-me.csv"
            master_path = root / "product_master.csv"

            pd.DataFrame([{
                "Order name": "#1767",
                "Created at": "2026-07-29 12:07:34",
                "Customer / Address / First name": "Chris",
                "Customer / Address / Last name": "Williams",
                "Billing address / Company": "Water",
                "Line item / Product title": "Georgia Boot CC5 Insole - GB00651",
                "Line item / Variant / Title": "ME 6–8.5",
                "Line item / Variant / Option 1": "ME 6–8.5",
                "Line item / Net quantity": "1",
                "Line item / Price": "29.41",
                "Line item / Discounts": "0",
                "Line item / Net sales": "29.41",
            }]).to_csv(csv_path, index=False)

            pd.DataFrame([{
                "Product Name": "Georgia Boot CC5 Insole",
                "Style Number": "GB00651",
                "Garment Color": "",
                "Vendor": "Georgia Boot",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Decoration Location": "",
                "Decoration Placement Instructions": "",
                "Decoration Color": "",
                "Product Category": "Boots",
                "Requires Size": "Yes",
                "Requires Color": "No",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            self.assertEqual(normalized.loc[0, "Variant Color"], "")
            self.assertEqual(normalized.loc[0, "Variant Size"], "M")

            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertEqual(parsed.loc[0, "Style Number"], "GB00651")
            self.assertEqual(parsed.loc[0, "Garment Color"], "")
            self.assertEqual(parsed.loc[0, "Size"], "M")

            detail, review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            self.assertEqual(detail.loc[0, "Review Status"], "Ready")
            self.assertTrue(review.empty)

    def test_large_abbreviation_insole_variant_is_ready_without_a_false_color(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "boot-sales-lr.csv"
            master_path = root / "product_master.csv"

            pd.DataFrame([{
                "Order name": "#1768",
                "Created at": "2026-07-29 12:07:34",
                "Customer / Address / First name": "Chris",
                "Customer / Address / Last name": "Williams",
                "Billing address / Company": "Water",
                "Line item / Product title": "Georgia Boot AMP Square Toe Insole - GB00651",
                "Line item / Variant / Title": "LR 9–10.5",
                "Line item / Variant / Option 1": "LR 9–10.5",
                "Line item / Net quantity": "1",
                "Line item / Price": "29.41",
                "Line item / Discounts": "0",
                "Line item / Net sales": "29.41",
            }]).to_csv(csv_path, index=False)

            pd.DataFrame([{
                "Product Name": "Georgia Boot AMP Square Toe Insole",
                "Style Number": "GB00651",
                "Garment Color": "",
                "Vendor": "Georgia Boot",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Decoration Location": "",
                "Decoration Placement Instructions": "",
                "Decoration Color": "",
                "Product Category": "Boots",
                "Requires Size": "Yes",
                "Requires Color": "No",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            self.assertEqual(normalized.loc[0, "Variant Color"], "")
            self.assertEqual(normalized.loc[0, "Variant Size"], "L")

            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertEqual(parsed.loc[0, "Style Number"], "GB00651")
            self.assertEqual(parsed.loc[0, "Garment Color"], "")
            self.assertEqual(parsed.loc[0, "Size"], "L")

            detail, review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            self.assertEqual(detail.loc[0, "Review Status"], "Ready")
            self.assertTrue(review.empty)

    def test_waterproof_boot_is_ready_with_its_numeric_size_and_no_decoration_prompt(self):
        """G5594 is a boot, not waterproof apparel awaiting decoration."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "muddog-boot-sale.csv"
            master_path = root / "product_master.csv"

            pd.DataFrame([{
                "Order name": "#1766",
                "Created at": "2026-07-29 12:07:34",
                "Customer / Address / First name": "Chris",
                "Customer / Address / Last name": "Williams",
                "Billing address / Company": "Water",
                "Line item / Product title": "Georgia Boot Muddog Steel Toe Waterproof Wellington - G5594",
                "Line item / Variant / Title": "11 / Regular",
                "Line item / Variant / Option 1": "11",
                "Line item / Variant / Option 2": "Regular",
                "Line item / Net quantity": "1",
                "Line item / Price": "205.88",
                "Line item / Discounts": "0",
                "Line item / Net sales": "205.88",
            }]).to_csv(csv_path, index=False)

            pd.DataFrame([{
                "Product Name": "Georgia Boot Muddog Steel Toe Waterproof Wellington",
                "Style Number": "G5594",
                "Garment Color": "",
                "Vendor": "Georgia Boot",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Decoration Location": "",
                "Decoration Placement Instructions": "",
                "Decoration Color": "",
                "Product Category": "Boots",
                "Requires Size": "Yes",
                "Requires Color": "No",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            self.assertEqual(normalized.loc[0, "Variant Color"], "")
            self.assertEqual(normalized.loc[0, "Variant Size"], "11 / Medium")

            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertEqual(parsed.loc[0, "Style Number"], "G5594")
            self.assertEqual(parsed.loc[0, "Garment Color"], "")
            self.assertEqual(parsed.loc[0, "Size"], "11 / Medium")

            detail, review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            self.assertEqual(detail.loc[0, "Product Category"], "Boots")
            self.assertEqual(detail.loc[0, "Size"], "11 / Medium")
            self.assertEqual(detail.loc[0, "Garment Color"], "")
            self.assertEqual(detail.loc[0, "Review Status"], "Ready")
            self.assertNotIn("Waterproof/rainwear", detail.loc[0, "Purchase Instructions"])
            self.assertTrue(review.empty)

    def test_regeneration_replaces_a_stale_regular_size_with_the_source_boot_size_and_width(self):
        """A pre-fix Review & Edit value must not undo 11 / Regular -> 11 / Medium."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "muddog-boot-sale.csv"
            master_path = root / "product_master.csv"
            prior_review_path = root / "prior-review.xlsx"

            pd.DataFrame([{
                "Order name": "#1766",
                "Created at": "2026-07-29 12:07:34",
                "Customer / Address / First name": "Chris",
                "Customer / Address / Last name": "Williams",
                "Billing address / Company": "Water",
                "Line item / Product title": "Georgia Boot Muddog Steel Toe Waterproof Wellington - G5594",
                "Line item / Variant / Title": "11 / Regular",
                "Line item / Variant / Option 1": "11",
                "Line item / Variant / Option 2": "Regular",
                "Line item / Net quantity": "1",
                "Line item / Price": "205.88",
                "Line item / Discounts": "0",
                "Line item / Net sales": "205.88",
            }]).to_csv(csv_path, index=False)
            pd.DataFrame([{
                "Product Name": "Georgia Boot Muddog Steel Toe Waterproof Wellington",
                "Style Number": "G5594",
                "Garment Color": "",
                "Vendor": "Georgia Boot",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Decoration Location": "",
                "Decoration Placement Instructions": "",
                "Decoration Color": "",
                "Product Category": "Boots",
                "Requires Size": "Yes",
                "Requires Color": "No",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            normalized = normalize_order_export(csv_path)
            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            detail, _review, _excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            stale_edit = detail.copy()
            stale_edit.loc[0, "Size"] = "REGULAR"
            stale_edit.loc[0, "Review Status"] = "Ready"
            with pd.ExcelWriter(prior_review_path, engine="xlsxwriter") as writer:
                stale_edit.to_excel(writer, sheet_name="Review & Edit", index=False)
                detail.to_excel(writer, sheet_name="All PO Lines", index=False)

            refreshed = _apply_previous_review_edits(detail, prior_review_path)
            self.assertEqual(refreshed.loc[0, "Size"], "11 / Medium")
            self.assertEqual(refreshed.loc[0, "Garment Color"], "")
            self.assertEqual(refreshed.loc[0, "Review Status"], "Ready")

    def test_boots_do_not_reach_decoration_color_audit_from_a_stale_label(self):
        """A legacy embroidery label cannot make a permanent Boots item auditable."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            master_path = Path(temporary_directory) / "product_master.csv"
            pd.DataFrame(columns=["Style Number", "Product Name"]).to_csv(master_path, index=False)
            audit = build_decoration_color_audit([{
                "Include": "Yes",
                "Product #": "G5594",
                "Description": "Georgia Boot Muddog Steel Toe Waterproof Wellington",
                "Product Category": "Boots",
                "Decoration Type": "Embroidery",
                "Decoration Location": "Left Chest",
                "Decoration Color": "White",
                "Garment Color": "",
                "Quantity": 1,
            }], master_path)
            self.assertEqual(audit, [])


if __name__ == "__main__":
    unittest.main()
