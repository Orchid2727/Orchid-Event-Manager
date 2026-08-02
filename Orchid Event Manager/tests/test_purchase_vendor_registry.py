from __future__ import annotations

import importlib
import os
import sys
import tempfile
import types
import unittest

import pandas as pd


class PurchaseVendorRegistryTests(unittest.TestCase):
    def test_added_rocky_boots_vendor_persists_and_is_canonical(self):
        """A custom vendor must not vanish when Product Master is reopened."""
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

                    # Upgrade an existing product assignment first; it must
                    # become a permanent option before the next editor open.
                    options = editor.purchase_vendor_options(pd.DataFrame({"Vendor": ["rocky boots"]}))
                    self.assertIn("Rocky Boots", editor.purchase_vendor_options())
                    self.assertEqual(editor.register_purchase_vendor("rocky-boots"), "Rocky Boots")
                    options = editor.purchase_vendor_options()

                    self.assertEqual(options.count("Rocky Boots"), 1)
                    self.assertIn("Rocky Boots", options)
                    self.assertTrue(editor.PURCHASE_VENDOR_REGISTRY_FILE.is_file())
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
