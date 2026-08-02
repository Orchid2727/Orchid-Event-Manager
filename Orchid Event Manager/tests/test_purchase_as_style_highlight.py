"""Regression coverage for the Product Master vendor-style mapping emphasis."""

from pathlib import Path
import unittest


EDITOR_SOURCE = (
    Path(__file__).resolve().parents[1] / "modules" / "product_master_editor.py"
)


class PurchaseAsStyleHighlightTests(unittest.TestCase):
    def test_mapping_has_a_clear_purpose_statement_above_the_color_rows(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn("Key Vendor Mapping: Use Purchase As Style #", source)
        self.assertIn("vendor order number differs from your Shopify style", source)

    def test_purchase_as_style_field_has_a_distinct_non_error_treatment(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn("KEY_VENDOR_MAPPING_BG", source)
        self.assertIn("KEY_VENDOR_MAPPING_BADGE_BG", source)
        self.assertIn("PURCHASE_AS_STYLE_COLUMN = 3", source)
        self.assertIn('text="Vendor order number"', source)
        self.assertIn("border_width=2, border_color=PURPLE", source)
        self.assertIn("fg_color=KEY_VENDOR_MAPPING_BG", source)

    def test_custom_vendor_style_badge_only_appears_when_a_number_is_entered(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn('text="Custom Vendor Style"', source)
        self.assertIn("def refresh_purchase_as_style_badge", source)
        self.assertIn("if normalize_space(purchasing_style_var.get()):", source)
        self.assertIn("purchase_as_style_badge.grid_remove()", source)
        self.assertIn('purchasing_style_var.trace_add("write", refresh_purchase_as_style_badge)', source)


if __name__ == "__main__":
    unittest.main()
