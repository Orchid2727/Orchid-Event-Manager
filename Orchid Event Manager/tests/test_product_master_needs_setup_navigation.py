"""Regression coverage for the one-click Product Master setup queue."""

from pathlib import Path
import sys
import unittest


APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

class ProductMasterNeedsSetupNavigationTests(unittest.TestCase):
    def test_dashboard_exposes_a_clickable_needs_setup_action(self):
        # Inspect the installed source rather than opening a GUI window in CI.
        text = (APP_ROOT / "modules" / "product_master_editor.py").read_text(encoding="utf-8")
        self.assertIn('"Needs Setup"', text)
        self.assertIn("command=self.open_needs_setup_queue", text)
        self.assertIn('text="Set Up Next  ➜"', text)

    def test_needs_setup_queue_filters_to_incomplete_styles_and_opens_first(self):
        text = (APP_ROOT / "modules" / "product_master_editor.py").read_text(encoding="utf-8")
        start = text.index("    def open_needs_setup_queue")
        end = text.index("    def dashboard_matching_style_keys", start)
        method = text[start:end]
        self.assertIn("self.incomplete_style_keys()", method)
        self.assertIn('self.catalog_filter_var.set("Needs Setup")', method)
        self.assertIn("self.show_editor(target_key=incomplete[0])", method)

    def test_needs_setup_filter_label_shows_the_current_count(self):
        text = (APP_ROOT / "modules" / "product_master_editor.py").read_text(encoding="utf-8")
        self.assertIn('"Needs Setup": f"Needs Setup ({incomplete})"', text)
        self.assertIn('self.filter_buttons["Needs Setup"].configure(text=f"Needs Setup ({incomplete})")', text)


if __name__ == "__main__":
    unittest.main()
