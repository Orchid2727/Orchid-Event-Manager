"""Regression checks for active-event Product Master routing safeguards."""

from __future__ import annotations

from copy import deepcopy
from hashlib import sha256
import json
from pathlib import Path
import tempfile
import unittest

import pandas as pd

from modules import packet_lock
from modules.outsource_rules import ROUTING_POLICY_VERSION


class PacketRoutingStatusTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.locked = self.root / "product_master_locked.csv"
        self.live = self.root / "product_master.csv"
        self.locked_overrides = self.root / "never_outsource_overrides_locked.json"
        self.live_overrides = self.root / "never_outsource_overrides.json"
        self.locked_overrides.write_text("{}\n", encoding="utf-8")
        self.live_overrides.write_text("{}\n", encoding="utf-8")
        self.base_master = [
            {
                "Product Name": "Port & Company Tee",
                "Style Number": "PC55T",
                "Garment Color": "Athletic Heather",
                "Vendor": "SanMar",
                "Decoration Type": "Screen Print",
                "Decoration Location": "Left Chest",
                "Decoration Color": "Black",
                "Never Outsource": "No",
                "Setup Required": "No",
            },
            {
                "Product Name": "Unrelated Polo",
                "Style Number": "UP100",
                "Garment Color": "Navy",
                "Vendor": "SanMar",
                "Decoration Type": "Embroidery",
                "Decoration Location": "Left Chest",
                "Decoration Color": "White",
                "Never Outsource": "No",
                "Setup Required": "No",
            },
        ]
        self._write_master(self.locked, self.base_master)
        self._write_master(self.live, self.base_master)
        self.manifest_path = self.root / "packet_lock_manifest.json"
        self._write_manifest()
        self.records = [{
            "Source ID": "line-1",
            "Include": "Yes",
            "Product #": "PC55T",
            "Description": "Port & Company Tee",
            "Garment Color": "Athletic Heather",
            "Purchase Vendor": "SanMar",
            "Decoration Type": "Screen Print",
            "Decoration Location": "Left Chest",
            "Do Not Outsource": "No",
        }]
        self.original_live_path = packet_lock.live_product_master_path
        packet_lock.live_product_master_path = lambda: self.live

    def tearDown(self):
        packet_lock.live_product_master_path = self.original_live_path
        self.temp.cleanup()

    def _write_master(self, path: Path, rows: list[dict]) -> None:
        pd.DataFrame(rows).to_csv(path, index=False)

    def _write_manifest(self) -> None:
        def entry(path: Path) -> dict:
            return {
                "name": path.name,
                "sha256": sha256(path.read_bytes()).hexdigest(),
                "size": path.stat().st_size,
            }

        manifest = {
            "schema_version": packet_lock.LOCK_SCHEMA_VERSION,
            "routing_policy_version": ROUTING_POLICY_VERSION,
            "files": {
                "product_master": entry(self.locked),
                "never_outsource_overrides": entry(self.locked_overrides),
            },
        }
        self.manifest_path.write_text(json.dumps(manifest), encoding="utf-8")

    def _status(self) -> dict:
        manifest = json.loads(self.manifest_path.read_text(encoding="utf-8"))
        return packet_lock._routing_status_from_manifest(self.records, self.manifest_path, manifest)

    def test_color_only_change_does_not_require_review_refresh(self):
        changed = deepcopy(self.base_master)
        changed[0]["Decoration Color"] = "White"
        self._write_master(self.live, changed)
        self.assertTrue(self._status()["current"])

    def test_unrelated_product_change_does_not_require_review_refresh(self):
        changed = deepcopy(self.base_master)
        changed[1]["Vendor"] = "S&S Activewear"
        self._write_master(self.live, changed)
        self.assertTrue(self._status()["current"])

    def test_vendor_change_for_active_default_route_requires_refresh(self):
        changed = deepcopy(self.base_master)
        changed[0]["Vendor"] = "S&S Activewear"
        self._write_master(self.live, changed)
        result = self._status()
        self.assertFalse(result["current"])
        self.assertIn("Purchase Vendor", result["changes"][0]["fields"])

    def test_manual_review_vendor_override_remains_valid(self):
        changed = deepcopy(self.base_master)
        changed[0]["Vendor"] = "S&S Activewear"
        self._write_master(self.live, changed)
        self.records[0]["Purchase Vendor"] = "Orchid Direct"
        self.assertTrue(self._status()["current"])

    def test_decoration_location_change_requires_refresh(self):
        changed = deepcopy(self.base_master)
        changed[0]["Decoration Location"] = "Left Sleeve"
        self._write_master(self.live, changed)
        result = self._status()
        self.assertFalse(result["current"])
        self.assertIn("Decoration Location", result["changes"][0]["fields"])

    def test_decoration_type_change_requires_refresh(self):
        changed = deepcopy(self.base_master)
        changed[0]["Decoration Type"] = "Embroidery"
        self._write_master(self.live, changed)
        result = self._status()
        self.assertFalse(result["current"])
        self.assertIn("Decoration Type", result["changes"][0]["fields"])

    def test_never_outsource_change_requires_refresh(self):
        changed = deepcopy(self.base_master)
        changed[0]["Never Outsource"] = "Yes"
        self._write_master(self.live, changed)
        result = self._status()
        self.assertFalse(result["current"])
        self.assertIn("Do Not Outsource", result["changes"][0]["fields"])


if __name__ == "__main__":
    unittest.main()
