from __future__ import annotations

import importlib
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


class ProductMasterLegacyCatalogRecoveryTests(unittest.TestCase):
    def test_editor_loader_repairs_a_catalog_missing_garment_color_before_ui_build(self):
        """A legacy catalog must reach the editor instead of crashing at launch.

        The actual failure occurred before Product Master could draw its first
        window.  This exercises its own loader against the same incomplete CSV
        shape while supplying only the tiny CustomTkinter class placeholder
        needed to import non-UI helper functions in the test environment.
        """
        fake_ctk = types.SimpleNamespace(CTk=type("CTk", (), {}))
        saved_ctk = sys.modules.get("customtkinter")
        saved_editor = sys.modules.pop("modules.product_master_editor", None)

        try:
            sys.modules["customtkinter"] = fake_ctk
            with tempfile.TemporaryDirectory() as temporary_directory:
                previous_data_dir = os.environ.get("ORCHID_DATA_DIR")
                os.environ["ORCHID_DATA_DIR"] = temporary_directory
                try:
                    editor = importlib.import_module("modules.product_master_editor")
                    product_master = Path(temporary_directory) / "product_master.csv"
                    product_master.write_text(
                        "Product Name,Style Number,Vendor\nGeorgia Boot,G5594,SanMar\n",
                        encoding="utf-8",
                    )

                    loaded = editor.load_product_master()

                    self.assertEqual(loaded.loc[0, "Garment Color"], "")
                    self.assertEqual(loaded.loc[0, "Style Number"], "G5594")
                    self.assertTrue(set(editor.COLUMNS).issubset(loaded.columns))
                    self.assertIn("Garment Color", product_master.read_text(encoding="utf-8").splitlines()[0])
                finally:
                    if previous_data_dir is None:
                        os.environ.pop("ORCHID_DATA_DIR", None)
                    else:
                        os.environ["ORCHID_DATA_DIR"] = previous_data_dir
        finally:
            sys.modules.pop("modules.product_master_editor", None)
            if saved_editor is not None:
                sys.modules["modules.product_master_editor"] = saved_editor
            if saved_ctk is None:
                sys.modules.pop("customtkinter", None)
            else:
                sys.modules["customtkinter"] = saved_ctk

    def test_editor_loader_discards_a_legacy_boot_size_saved_as_a_purchasing_color(self):
        """The exact A5M2T problem must disappear without a manual cleanup."""
        fake_ctk = types.SimpleNamespace(CTk=type("CTk", (), {}))
        saved_ctk = sys.modules.get("customtkinter")
        saved_editor = sys.modules.pop("modules.product_master_editor", None)

        try:
            sys.modules["customtkinter"] = fake_ctk
            with tempfile.TemporaryDirectory() as temporary_directory:
                previous_data_dir = os.environ.get("ORCHID_DATA_DIR")
                os.environ["ORCHID_DATA_DIR"] = temporary_directory
                try:
                    editor = importlib.import_module("modules.product_master_editor")
                    product_master = Path(temporary_directory) / "product_master.csv"
                    product_master.write_text(
                        "Product Name,Style Number,Garment Color,Vendor,Decoration Type,Product Category,Requires Size,Requires Color,Requires Decoration,Never Outsource,Setup Required\n"
                        "Men's TITAN EV 6 in Composite Toe Work Boot,A5M2T,15,Orchid Uniforms & Apparel,Blank Garment (No Decoration),Boots,Yes,No,No,Yes,Yes\n",
                        encoding="utf-8",
                    )

                    loaded = editor.load_product_master()

                    self.assertEqual(loaded.loc[0, "Garment Color"], "")
                    self.assertEqual(loaded.loc[0, "Setup Required"], "No")
                    self.assertEqual(loaded.loc[0, "Requires Size"], "Yes")
                    self.assertEqual(loaded.loc[0, "Requires Color"], "No")
                    self.assertEqual(loaded.loc[0, "Requires Decoration"], "No")
                    self.assertEqual(loaded.loc[0, "Never Outsource"], "Yes")
                finally:
                    if previous_data_dir is None:
                        os.environ.pop("ORCHID_DATA_DIR", None)
                    else:
                        os.environ["ORCHID_DATA_DIR"] = previous_data_dir
        finally:
            sys.modules.pop("modules.product_master_editor", None)
            if saved_editor is not None:
                sys.modules["modules.product_master_editor"] = saved_editor
            if saved_ctk is None:
                sys.modules.pop("customtkinter", None)
            else:
                sys.modules["customtkinter"] = saved_ctk


if __name__ == "__main__":
    unittest.main()
