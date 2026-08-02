"""Regression checks for Orchid's no-purchase Final Sale Boot checkout item."""

from pathlib import Path
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.employee_totals import load_employee_totals
from modules.internal_services import is_final_sale_accounting_product
from modules.report_modes import UNIFORM_SIZING_EVENT
from modules.review_workbook import _build_review_data, generate_review_workbook
from modules.shopify_parser import normalize_order_export, parse_shopify_orders


class FinalSaleBootTests(unittest.TestCase):
    def _write_final_sale_csv(self, path: Path) -> None:
        pd.DataFrame([{
            "Order name": "#1757",
            "Created at": "2026-07-29 20:15:00",
            "Customer / Address / First name": "Ricardo",
            "Customer / Address / Last name": "Rodriguez",
            "Billing address / Company": "Line",
            "Line item / Product title": "Final Sale Boot",
            "Line item / Variant / Title": "12",
            "Line item / Variant / Option 1": "12",
            "Line item / Net quantity": "1",
            "Line item / Price": "50.00",
            "Line item / Discounts": "0",
            "Line item / Net sales": "50.00",
        }]).to_csv(path, index=False)

    @staticmethod
    def _write_empty_master(path: Path) -> None:
        pd.DataFrame(columns=[
            "Product Name", "Style Number", "Garment Color", "Vendor",
            "Decoration Type", "Decoration Location", "Decoration Color",
            "Product Category", "Requires Size", "Requires Color",
            "Requires Decoration", "Never Outsource",
        ]).to_csv(path, index=False)

    def test_exact_final_sale_boot_is_an_accounting_line_not_a_boot_to_purchase(self):
        self.assertTrue(is_final_sale_accounting_product("Final Sale Boot"))
        self.assertTrue(is_final_sale_accounting_product("final sale boots"))
        self.assertFalse(is_final_sale_accounting_product("Georgia Boot Final Sale"))

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "final-sale.csv"
            master_path = root / "product_master.csv"
            self._write_final_sale_csv(csv_path)
            self._write_empty_master(master_path)

            normalized = normalize_order_export(csv_path)
            parsed = parse_shopify_orders(csv_path, master_path, normalized_orders=normalized)
            self.assertTrue(parsed.empty)

            detail, review, excluded, _raw = _build_review_data(
                csv_path, master_path, normalized_orders=normalized, parsed_orders=parsed
            )
            self.assertTrue(detail.empty)
            self.assertTrue(review.empty)
            self.assertTrue(excluded.loc[0, "Service / Fee"].startswith("Final Sale Boot"))

    def test_final_sale_boot_remains_in_employee_totals_without_purchase_review(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            csv_path = root / "final-sale.csv"
            master_path = root / "product_master.csv"
            reports_root = root / "reports"
            self._write_final_sale_csv(csv_path)
            self._write_empty_master(master_path)

            with patch.dict(os.environ, {
                "ORCHID_ALLOW_TEST_MASTER": "1",
                "ORCHID_DATA_DIR": str(root / "application-data"),
            }):
                result = generate_review_workbook(
                    csv_path,
                    master_path,
                    reports_root,
                    report_mode=UNIFORM_SIZING_EVENT,
                    event_name="Boot Sale Test",
                )
            self.assertEqual(result["lines"], 0)
            self.assertEqual(result["review_lines"], 0)

            totals = load_employee_totals(result["output_path"])
            self.assertEqual(len(totals), 1)
            self.assertEqual(totals[0]["order_numbers"], "#1757")
            self.assertEqual(totals[0]["order_total"], 50.0)


if __name__ == "__main__":
    unittest.main()
