"""Regression checks for active-event isolation."""

from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from modules.active_packet import (
    active_workbook_matches_state,
    bind_active_workbook,
    new_active_packet_state,
    source_signature,
)


class ActivePacketIdentityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.boot_csv = self.root / "Line Boots 2026.csv"
        self.old_csv = self.root / "Earlier store orders.csv"
        self.boot_csv.write_text("Order,Name\n1757,Ricardo Rodriguez\n", encoding="utf-8")
        self.old_csv.write_text("Order,Name\n1747,James Smith\n1762,Ricardo Rodriguez\n", encoding="utf-8")

    def tearDown(self):
        self.temp.cleanup()

    def _make_workbook(self, label: str, source_csv: Path, active_packet_id: str) -> Path:
        lock_dir = self.root / f"lock-{label}"
        lock_dir.mkdir()
        locked_source = lock_dir / "source_orders_locked.csv"
        locked_source.write_bytes(source_csv.read_bytes())
        manifest = {
            "files": {
                "source": {"sha256": source_signature(locked_source)},
            }
        }
        manifest_path = lock_dir / "packet_lock_manifest.json"
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        manifest_hash = sha256(manifest_path.read_bytes()).hexdigest()

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "System Info"
        sheet.append(["Active Packet ID", active_packet_id])
        sheet.append(["Packet Lock ID", f"LOCK-{label}"])
        sheet.append(["Packet Lock Manifest Path", str(manifest_path)])
        sheet.append(["Packet Lock Manifest SHA256", manifest_hash])
        path = self.root / f"{label}.xlsx"
        workbook.save(path)
        return path

    def test_replacement_packet_cannot_bind_an_older_workbook(self):
        state = new_active_packet_state(self.boot_csv)
        old_workbook = self._make_workbook("old", self.old_csv, "EVENT-OLDER")
        with self.assertRaisesRegex(ValueError, "different event packet|different CSV"):
            bind_active_workbook(state, old_workbook)

    def test_bound_boot_event_keeps_valid_order_and_excludes_old_orders(self):
        state = new_active_packet_state(self.boot_csv)
        boot_workbook = self._make_workbook("boot", self.boot_csv, state["active_packet_id"])
        old_workbook = self._make_workbook("old", self.old_csv, "EVENT-OLDER")
        state = bind_active_workbook(state, boot_workbook)
        valid, reason = active_workbook_matches_state(state, boot_workbook)
        self.assertTrue(valid, reason)
        stale_valid, stale_reason = active_workbook_matches_state(state, old_workbook)
        self.assertFalse(stale_valid)
        self.assertIn("different event", stale_reason)

    def test_changed_csv_blocks_reports_instead_of_reusing_old_packet(self):
        state = new_active_packet_state(self.boot_csv)
        boot_workbook = self._make_workbook("boot", self.boot_csv, state["active_packet_id"])
        state = bind_active_workbook(state, boot_workbook)
        self.boot_csv.write_text("Order,Name\n1757,Ricardo Rodriguez\n9999,Changed\n", encoding="utf-8")
        valid, reason = active_workbook_matches_state(state, boot_workbook)
        self.assertFalse(valid)
        self.assertIn("changed", reason)

    def test_same_csv_from_an_earlier_event_is_blocked(self):
        current = new_active_packet_state(self.boot_csv)
        earlier_workbook = self._make_workbook("earlier", self.boot_csv, "EVENT-EARLIER")
        with self.assertRaisesRegex(ValueError, "different event packet"):
            bind_active_workbook(current, earlier_workbook)

    def test_app_uses_explicit_active_workbook_not_newest_file(self):
        app_source = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
        discovery = app_source.split("    def discover_existing_work", 1)[1].split("    def _active_review_needs", 1)[0]
        self.assertIn("active_review_workbook", discovery)
        self.assertIn("bind_active_workbook", discovery)
        self.assertNotIn("max(review_files", discovery)


if __name__ == "__main__":
    unittest.main()
