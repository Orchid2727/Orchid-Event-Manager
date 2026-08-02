from __future__ import annotations

import importlib
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch


class ProductMasterSnapshotCacheTests(unittest.TestCase):
    def test_master_save_invalidates_an_old_dashboard_snapshot(self):
        """Completed Product Master work must replace stale workbook notices."""
        fake_ctk = types.SimpleNamespace(CTk=type("CTk", (), {}))
        saved_ctk = sys.modules.get("customtkinter")
        saved_app = sys.modules.pop("app", None)
        saved_data_dir = os.environ.get("ORCHID_DATA_DIR")

        try:
            sys.modules["customtkinter"] = fake_ctk
            with tempfile.TemporaryDirectory() as temporary_directory:
                root = Path(temporary_directory)
                os.environ["ORCHID_DATA_DIR"] = str(root / "orchid-data")
                app = importlib.import_module("app")
                workbook = root / "review.xlsx"
                master = root / "product_master.csv"
                workbook.write_bytes(b"placeholder")
                master.write_text("first", encoding="utf-8")

                manager = object.__new__(app.OrchidPurchaseManager)
                manager.last_review_workbook = workbook
                manager._mission_snapshot_cache_key = None
                manager._mission_snapshot_cache = {}
                manager._overlay_completed_review_state = lambda snapshot, _path: snapshot
                calls = []

                def load_snapshot(*_args, **_kwargs):
                    calls.append("loaded")
                    return {"issues": [], "load": len(calls)}

                with patch.object(app, "live_product_master_path", return_value=master), patch.object(
                    app, "load_mission_control_snapshot", side_effect=load_snapshot
                ):
                    first = app.OrchidPurchaseManager._mission_snapshot(manager)
                    master.write_text("updated saved catalog", encoding="utf-8")
                    second = app.OrchidPurchaseManager._mission_snapshot(manager)

                self.assertEqual(first["load"], 1)
                self.assertEqual(second["load"], 2)
                self.assertEqual(calls, ["loaded", "loaded"])
        finally:
            sys.modules.pop("app", None)
            if saved_app is not None:
                sys.modules["app"] = saved_app
            if saved_ctk is None:
                sys.modules.pop("customtkinter", None)
            else:
                sys.modules["customtkinter"] = saved_ctk
            if saved_data_dir is None:
                os.environ.pop("ORCHID_DATA_DIR", None)
            else:
                os.environ["ORCHID_DATA_DIR"] = saved_data_dir


if __name__ == "__main__":
    unittest.main()
