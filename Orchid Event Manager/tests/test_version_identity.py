from pathlib import Path
import unittest


APP_ROOT = Path(__file__).resolve().parents[1]
CONTENTS_ROOT = APP_ROOT.parents[1]
EXPECTED = "4.9.104"


class VersionIdentityTests(unittest.TestCase):
    def test_runtime_visible_version_is_consistent(self):
        app_source = (APP_ROOT / "app.py").read_text(encoding="utf-8")
        bootstrap = (CONTENTS_ROOT / "Resources" / "bootstrap.py").read_text(encoding="utf-8")
        launcher = (CONTENTS_ROOT / "MacOS" / "OrchidPurchaseManager").read_text(encoding="utf-8")
        info_plist = (CONTENTS_ROOT / "Info.plist").read_text(encoding="utf-8")
        self.assertEqual((APP_ROOT / "VERSION.txt").read_text(encoding="utf-8").strip(), EXPECTED)
        self.assertIn(f"Professional {EXPECTED}", app_source)
        self.assertIn(f"v {EXPECTED}", app_source)
        self.assertIn(f"Version {EXPECTED}", app_source)
        self.assertIn(f"Professional {EXPECTED}", bootstrap)
        self.assertIn(f"Manager {EXPECTED}", launcher)
        self.assertIn(f"<string>{EXPECTED}</string>", info_plist)


if __name__ == "__main__":
    unittest.main()
