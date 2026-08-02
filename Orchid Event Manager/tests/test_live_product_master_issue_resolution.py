from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from modules.mission_control import _filter_live_product_master_issues


class LiveProductMasterIssueResolutionTests(unittest.TestCase):
    def test_completed_rocky_boot_removes_stale_workbook_setup_notice(self):
        """The v4.9.73 dashboard screenshot must be impossible after a refresh."""
        with tempfile.TemporaryDirectory() as temporary_directory:
            master_path = Path(temporary_directory) / "product_master.csv"
            pd.DataFrame([{
                "Product Name": "Rocky Worksmart Waterproof Composite Toe Pull-On Work Boot",
                "Style Number": "RKK0402",
                "Garment Color": "",
                "Vendor": "Rocky Boots",
                "Decoration Type": "Blank Garment (No Decoration)",
                "Decoration Location": "",
                "Decoration Color": "",
                "Product Category": "Boots",
                "Requires Size": "Yes",
                "Requires Color": "No",
                "Requires Decoration": "No",
                "Never Outsource": "Yes",
                "Setup Required": "No",
            }]).to_csv(master_path, index=False)

            stale_issue = {
                "line_id": "line-rkk0402",
                "affected_line_ids": "line-rkk0402",
                "product": "RKK0402",
                "description": "Rocky Worksmart Waterproof Composite Toe Pull-On Work Boot",
                "vendor": "",
                "garment_colors": "",
                "reason": "Product Master setup required",
                "source": "Product Master",
                "fix_in": "Product Master",
            }

            remaining, resolved_line_ids = _filter_live_product_master_issues(
                [stale_issue], master_path
            )

            self.assertEqual(remaining, [])
            self.assertEqual(resolved_line_ids, {"line-rkk0402"})


if __name__ == "__main__":
    unittest.main()
