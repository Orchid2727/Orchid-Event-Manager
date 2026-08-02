from __future__ import annotations

from pathlib import Path
import os
import unittest


APP_ROOT = Path(__file__).resolve().parents[1]
CONTENTS_ROOT = APP_ROOT.parents[1]
MAIN_INFO = CONTENTS_ROOT / "Info.plist"
COMPANION_APP = CONTENTS_ROOT / "Resources" / "Orchid Product Master.app"
COMPANION_INFO = COMPANION_APP / "Contents" / "Info.plist"
COMPANION_LAUNCHER = COMPANION_APP / "Contents" / "MacOS" / "OrchidProductMaster"


class ProductMasterCompanionAppTests(unittest.TestCase):
    def test_dedicated_app_is_present_and_executable(self):
        self.assertTrue(COMPANION_INFO.is_file())
        self.assertTrue(COMPANION_LAUNCHER.is_file())
        self.assertTrue(os.access(COMPANION_LAUNCHER, os.X_OK))

    def test_companion_has_its_own_bundle_identity(self):
        main_info = MAIN_INFO.read_text(encoding="utf-8")
        companion_info = COMPANION_INFO.read_text(encoding="utf-8")
        self.assertIn("Orchid Product Master", companion_info)
        self.assertIn("com.orchiduniforms.purchasemanager.productmaster.4974", companion_info)
        self.assertNotIn("com.orchiduniforms.purchasemanager.productmaster.4974", main_info)

    def test_companion_starts_the_shared_product_master_entrypoint(self):
        launcher = COMPANION_LAUNCHER.read_text(encoding="utf-8")
        self.assertIn('EDITOR_CONTENTS_DIR="$(cd "$(dirname "$0")/.." && pwd)"', launcher)
        self.assertIn('MAIN_RESOURCES_DIR="$(cd "$EDITOR_CONTENTS_DIR/../.." && pwd)"', launcher)
        self.assertIn('"$MAIN_RESOURCES_DIR/bootstrap.py" "$APP_ROOT" --product-master', launcher)


if __name__ == "__main__":
    unittest.main()
