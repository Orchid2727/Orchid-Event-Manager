from __future__ import annotations

import unittest

import pandas as pd

from modules.product_candidate_sync import COLUMNS
from modules.color_setup_guidance import (
    new_color_setup_guidance,
    new_color_setup_message,
    prioritize_new_color_setup_rows,
)


def color_row(color: str, *, setup_required: str = "No", decoration_color: str = "Black") -> dict[str, str]:
    row = {column: "" for column in COLUMNS}
    row.update({
        "Product Name": "Gildan DryBlend 50 Cotton/50 Poly T-Shirt",
        "Style Number": "8000",
        "Garment Color": color,
        "Vendor": "SanMar",
        "Decoration Type": "Screen Print",
        "Decoration Location": "Left Chest",
        "Decoration Color": decoration_color,
        "Product Category": "T-Shirts",
        "Requires Size": "Yes",
        "Requires Color": "Yes",
        "Requires Decoration": "Yes",
        "Setup Required": setup_required,
    })
    return row


class ProductMasterNewColorGuidanceTests(unittest.TestCase):
    def test_one_new_color_is_the_only_thread_ink_decision(self):
        rows = pd.DataFrame(
            [color_row(f"Saved Color {number:02d}") for number in range(1, 34)] + [
                color_row("Tennessee Orange", setup_required="Yes", decoration_color=""),
            ],
            columns=COLUMNS,
        )

        guidance = new_color_setup_guidance(rows)
        message, needs_attention = new_color_setup_message(rows)

        self.assertEqual(guidance["new_count"], 1)
        self.assertEqual(guidance["existing_count"], 33)
        self.assertEqual(guidance["new_color_names"], ["Tennessee Orange"])
        self.assertEqual(guidance["thread_ink_color_names"], ["Tennessee Orange"])
        self.assertTrue(needs_attention)
        self.assertEqual(
            message,
            "1 NEW COLOR NEEDS THREAD / INK: Tennessee Orange  •  33 saved color settings unchanged",
        )

    def test_new_color_is_displayed_first_without_reordering_saved_colors(self):
        rows = pd.DataFrame(
            [
                color_row("White"),
                color_row("Black"),
                color_row("Navy"),
                color_row("Safety Green", setup_required="Yes", decoration_color=""),
            ],
            columns=COLUMNS,
        )

        ordered = prioritize_new_color_setup_rows(rows)

        self.assertEqual(ordered["Garment Color"].tolist(), ["Safety Green", "Black", "Navy", "White"])

    def test_blank_placeholder_is_not_presented_as_a_new_color(self):
        rows = pd.DataFrame(
            [
                color_row("Black"),
                color_row("", setup_required="Yes", decoration_color=""),
            ],
            columns=COLUMNS,
        )

        guidance = new_color_setup_guidance(rows)
        message, needs_attention = new_color_setup_message(rows)

        self.assertEqual(guidance["new_count"], 0)
        self.assertEqual(guidance["existing_count"], 1)
        self.assertFalse(needs_attention)
        self.assertEqual(message, "All saved purchasing colors and Thread / Ink settings are retained below.")

    def test_blank_service_color_row_does_not_break_post_save_progress_refresh(self):
        # Reproduces BAGDECORATION and the other internal-service records. They
        # are allowed to have no purchasing color, but the old empty-frame path
        # raised KeyError("Decoration Type") after Save Changes.
        rows = pd.DataFrame([
            {
                "Product Name": "Bag Decoration Service",
                "Style Number": "BAGDECORATION",
                "Garment Color": "",
                "Vendor": "Orchid Uniforms & Apparel",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Setup Required": "No",
            }
        ])

        guidance = new_color_setup_guidance(rows)
        message, needs_attention = new_color_setup_message(rows)

        self.assertEqual(guidance["new_count"], 0)
        self.assertEqual(guidance["existing_count"], 0)
        self.assertFalse(needs_attention)
        self.assertEqual(message, "All saved purchasing colors and Thread / Ink settings are retained below.")

    def test_legacy_rows_without_color_columns_do_not_block_product_master(self):
        # Reproduces the launch failure from a legacy catalog whose rows were
        # loaded without a Garment Color column.  Guidance is optional and must
        # degrade to an empty list rather than crash the whole editor.
        rows = pd.DataFrame([{"Style Number": "G5594", "Product Name": "Georgia Boot"}])

        guidance = new_color_setup_guidance(rows)
        message, needs_attention = new_color_setup_message(rows)
        ordered = prioritize_new_color_setup_rows(rows)

        self.assertEqual(guidance["new_count"], 0)
        self.assertEqual(guidance["new_color_names"], [])
        self.assertEqual(guidance["thread_ink_color_names"], [])
        self.assertFalse(needs_attention)
        self.assertEqual(message, "All saved purchasing colors and Thread / Ink settings are retained below.")
        self.assertIn("Garment Color", ordered.columns)
        self.assertEqual(ordered["Garment Color"].tolist(), [""])


if __name__ == "__main__":
    unittest.main()
