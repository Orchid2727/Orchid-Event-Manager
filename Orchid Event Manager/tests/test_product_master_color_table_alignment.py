"""Regression coverage for the Product Master full-width color-mapping layout."""

from pathlib import Path
import unittest


EDITOR_SOURCE = (
    Path(__file__).resolve().parents[1] / "modules" / "product_master_editor.py"
)


class ProductMasterColorTableAlignmentTests(unittest.TestCase):
    def test_editor_reclaims_the_inline_catalog_width(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn("workspace.grid(row=0, column=0, sticky=\"nsew\")", source)
        self.assertIn("workspace.grid_columnconfigure(1, weight=62, minsize=860)", source)
        self.assertNotIn("self.build_editor_catalog(content)", source)

    def test_color_mapping_is_a_single_shared_left_to_right_row(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn("COLOR_TABLE_COLUMNS = (", source)
        self.assertIn("configure_color_table_columns(frame)", source)
        self.assertNotIn("configure_color_card_columns(frame)", source)
        for column, heading in enumerate((
            "Purchasing Color",
            "Vendor Color Code",
            "Shopify Color Alias",
            "Purchase As Style #",
            "Thread / Ink",
        )):
            self.assertIn(f"column={column}", source)
            self.assertIn(f'("{heading}",', source)

    def test_delete_remains_visible_in_the_last_shared_column(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn('text="Actions"', source)
        self.assertIn("column=len(COLOR_TABLE_COLUMNS)", source)
        self.assertIn('text="Clear" if is_boot_base_row else "Delete"', source)

    def test_thread_ink_dropdown_is_on_the_same_row_as_the_other_fields(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn("decoration_combo.grid(row=1, column=4", source)

    def test_color_workspace_uses_the_remaining_editor_height(self):
        source = EDITOR_SOURCE.read_text(encoding="utf-8")

        self.assertIn("panel.grid_rowconfigure(3, weight=1)", source)
        self.assertIn("self.colors_scroll.grid(row=3", source)
        self.assertNotIn("panel.grid_rowconfigure(4, weight=1)", source)


if __name__ == "__main__":
    unittest.main()
