from pathlib import Path
import sys
import unittest


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

import pandas as pd

from modules.outsource_rules import (
    apply_never_outsource_defaults,
    default_never_outsource,
    resolve_never_outsource,
)
from modules.purchase_rules import (
    CATEGORY_OPTIONS,
    apply_purchase_rule_defaults,
    category_defaults,
    infer_category,
    row_rules,
)


class BootsCategoryTests(unittest.TestCase):
    def test_boots_is_a_permanent_product_master_category(self):
        self.assertIn("Boots", CATEGORY_OPTIONS)

    def test_boot_and_insole_descriptions_default_to_boots(self):
        self.assertEqual(infer_category("Georgia Boot CC5 Insole", "GB00111"), "Boots")
        self.assertEqual(infer_category("Georgia Boots Waterproof Work Boot", "GB00021"), "Boots")

    def test_boots_always_use_the_permanent_purchasing_rules(self):
        self.assertEqual(
            category_defaults("Boots"),
            {"requires_size": True, "requires_color": False, "requires_decoration": False},
        )
        self.assertTrue(default_never_outsource(category="Boots"))
        self.assertTrue(resolve_never_outsource("No", category="Boots"))

        rules = row_rules({
            "Product Name": "Georgia Boot CC5 Insole",
            "Style Number": "GB00111",
            "Product Category": "Boots",
            "Requires Size": "No",
            "Requires Color": "Yes",
            "Requires Decoration": "Yes",
            "Decoration Type": "",
        })
        self.assertEqual(rules["Product Category"], "Boots")
        self.assertEqual(rules["Requires Size"], "Yes")
        self.assertEqual(rules["Requires Color"], "No")
        self.assertEqual(rules["Requires Decoration"], "No")

    def test_existing_boot_records_are_repaired_even_if_saved_with_old_values(self):
        frame = pd.DataFrame([{
            "Product Name": "Georgia Boots Waterproof Work Boot",
            "Style Number": "GB00021",
            "Product Category": "Boots",
            "Requires Size": "No",
            "Requires Color": "Yes",
            "Requires Decoration": "Yes",
            "Decoration Type": "Embroidery",
            "Decoration Color": "Black",
            "Decoration Location": "Left Chest",
            "Never Outsource": "No",
        }])
        repaired = apply_never_outsource_defaults(apply_purchase_rule_defaults(frame))
        self.assertEqual(repaired.at[0, "Requires Size"], "Yes")
        self.assertEqual(repaired.at[0, "Requires Color"], "No")
        self.assertEqual(repaired.at[0, "Requires Decoration"], "No")
        self.assertEqual(repaired.at[0, "Decoration Type"], "Blank Garment (No Decoration)")
        self.assertEqual(repaired.at[0, "Decoration Color"], "")
        self.assertEqual(repaired.at[0, "Decoration Location"], "")
        self.assertEqual(repaired.at[0, "Never Outsource"], "Yes")


if __name__ == "__main__":
    unittest.main()
