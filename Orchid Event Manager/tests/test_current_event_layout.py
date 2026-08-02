"""Regression checks for the Current Event hero layout."""

from __future__ import annotations

from pathlib import Path
import unittest


class CurrentEventLayoutTests(unittest.TestCase):
    def setUp(self):
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        source = app_path.read_text(encoding="utf-8")
        start = source.index("    def build_import_page(self):")
        end = source.index("    def _refresh_current_event_job_logo", start)
        self.current_event_builder = source[start:end]

    def test_redundant_import_icon_is_not_built(self):
        self.assertNotIn("current_event_icon", self.current_event_builder)
        self.assertNotIn("workflow_import_purple.png", self.current_event_builder)

    def test_job_name_and_logo_lead_the_current_event_page(self):
        self.assertIn("self.current_event_title.grid(row=1", self.current_event_builder)
        self.assertIn("self.current_event_job_logo.grid(row=2", self.current_event_builder)
        self.assertIn("buttons.grid(row=6", self.current_event_builder)

    def test_empty_state_is_one_centered_composition_without_the_glow_divider(self):
        self.assertIn("def _show_current_event_empty_layout", self.current_event_builder)
        self.assertIn("self.current_event_glow.place_forget()", self.current_event_builder)
        self.assertIn('self.current_event_split.grid_anchor("center")', self.current_event_builder)
        self.assertIn("self.current_event_illustration_card.grid()", self.current_event_builder)
        self.assertIn("self.current_event_card.grid()", self.current_event_builder)
        self.assertNotIn("_reflow_empty_current_event_hero", self.current_event_builder)

    def test_active_packet_restores_the_logo_ready_two_column_layout(self):
        self.assertIn("def _show_current_event_active_layout", self.current_event_builder)
        self.assertIn("self.current_event_illustration_card.grid()", self.current_event_builder)
        self.assertIn("self.current_event_card.grid()", self.current_event_builder)
        self.assertIn("self.current_event_glow.place(relx=0.52", self.current_event_builder)

    def test_start_new_event_has_its_own_full_width_row(self):
        """Starting a new event is intentionally separate from Add Orders."""
        self.assertIn('buttons, text="Start New Event"', self.current_event_builder)
        self.assertIn(
            'self.current_event_replace.grid(row=1, column=0, columnspan=2, sticky="ew"',
            self.current_event_builder,
        )
        self.assertIn(
            'self.current_event_regenerate.grid(row=2, column=0, columnspan=2, sticky="ew"',
            self.current_event_builder,
        )

    def test_finish_and_clear_replaces_archive_workflow(self):
        self.assertIn('text="Finish & Clear Event"', self.current_event_builder)
        self.assertIn('command=self.finish_and_clear_current_event', self.current_event_builder)
        self.assertNotIn('text="Finish / Archive Event"', self.current_event_builder)

    def test_active_packet_refresh_keeps_new_event_and_review_on_separate_rows(self):
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        source = app_path.read_text(encoding="utf-8")
        start = source.index("    def refresh_current_event_page")
        end = source.index("    @staticmethod\n    def _preview_order_numbers", start)
        refresh = source[start:end]
        self.assertIn("self._show_current_event_active_layout()", refresh)
        self.assertIn("self._show_current_event_empty_layout()", refresh)
        self.assertIn(
            'self.current_event_replace.grid(row=1, column=0, columnspan=2, sticky="ew"',
            refresh,
        )
        self.assertIn(
            'self.current_event_regenerate.grid(row=2, column=0, columnspan=2, sticky="ew"',
            refresh,
        )

    def test_archive_is_not_a_navigation_destination_or_report_fallback(self):
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        source = app_path.read_text(encoding="utf-8")
        nav_start = source.index("        for row, page, label in [", source.index("        divider = ctk.CTkFrame"))
        nav_end = source.index("        self.page_container.grid_columnconfigure", nav_start)
        navigation = source[nav_start:nav_end]
        self.assertNotIn('"archive"', navigation)

        release_start = source.index("    def _purchase_order_release_location")
        release_end = source.index("    def _archived_source_purchase_order_location", release_start)
        release_method = source[release_start:release_end]
        self.assertNotIn("_archive_release_roots", release_method)
        self.assertNotIn("_verified_complete_archive_purchase_orders", release_method)

        generation_start = source.index("    def _run_final_purchase_order_generation")
        generation_end = source.index("    def create_final_purchase_orders", generation_start)
        generation = source[generation_start:generation_end]
        self.assertNotIn("archive_completed_event(", generation)


if __name__ == "__main__":
    unittest.main()
