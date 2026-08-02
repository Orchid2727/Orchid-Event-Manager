from __future__ import annotations

import os
import sys
import tempfile
import types
import unittest
from pathlib import Path

import pandas as pd


# These tests exercise pure catalog-cleanup helpers.  Supplying the base class
# keeps them runnable in a headless release-validation environment without
# changing the desktop application's actual CustomTkinter dependency.
if "customtkinter" not in sys.modules:
    sys.modules["customtkinter"] = types.SimpleNamespace(CTk=type("CTk", (), {}))

# ``product_master_editor`` initializes its Product Master path at import time.
# Keep this helper-only test outside a real macOS Application Support folder.
os.environ.setdefault(
    "ORCHID_DATA_DIR",
    str(Path(tempfile.gettempdir()) / "orchid-product-master-fast-save-tests"),
)

from modules.product_master_editor import (
    COLUMNS,
    clean_and_deduplicate_master,
    clean_changed_style_groups,
)


def row(style: str, color: str, *, vendor: str = "SanMar") -> dict:
    value = {column: "" for column in COLUMNS}
    value.update(
        {
            "Product Name": f"Product {style}",
            "Style Number": style,
            "Garment Color": color,
            "Vendor": vendor,
            "Decoration Type": "Embroidery",
            "Decoration Location": "Left Chest",
            "Decoration Color": "Black",
            "Product Category": "Polos",
            "Requires Size": "Yes",
            "Requires Color": "Yes",
            "Requires Decoration": "Yes",
            "Never Outsource": "No",
            "Setup Required": "No",
        }
    )
    return value


class ProductMasterFastSaveTests(unittest.TestCase):
    def test_changed_style_cleanup_matches_full_cleanup(self):
        master = clean_and_deduplicate_master(pd.DataFrame(
            [
                row("K500", "Navy"),
                row("CTA18", "Black"),
                row("CTA18", "Charcoal"),
            ],
            columns=COLUMNS,
        ))
        # The live editor starts from this cleaned state; emulate a duplicate
        # color row introduced while editing K500 before the next save.
        master = pd.concat([master, master[master["Style Number"].eq("K500")]], ignore_index=True)

        expected = clean_and_deduplicate_master(master)
        actual = clean_changed_style_groups(master, ["style:K500"])

        self.assertTrue(actual.equals(expected))

    def test_style_rename_cleans_the_destination_group(self):
        master = clean_and_deduplicate_master(pd.DataFrame(
            [
                row("K500", "Navy"),
                row("CTA18", "Black"),
                row("CTA18", "Black"),
            ],
            columns=COLUMNS,
        ))
        # Simulate changing K500's product number to an existing CTA18 group.
        master.loc[master["Style Number"].eq("K500"), "Style Number"] = "CTA18"

        expected = clean_and_deduplicate_master(master)
        actual = clean_changed_style_groups(master, ["style:K500", "style:CTA18"])

        self.assertTrue(actual.equals(expected))


if __name__ == "__main__":
    unittest.main()
