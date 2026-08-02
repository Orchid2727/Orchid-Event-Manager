from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from modules.product_master_launch import (
    containing_app_bundle,
    dedicated_product_master_bundle,
    product_master_launch_command,
)


class ProductMasterLaunchTests(unittest.TestCase):
    def test_macos_bundle_launches_the_editor_directly_with_the_live_runtime(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            bundle = Path(temporary_directory) / "Orchid Purchase Manager.app"
            executable = bundle / "Contents" / "Resources" / "Orchid Product Master.app" / "Contents" / "MacOS" / "OrchidProductMaster"
            app_file = bundle / "Contents" / "Resources" / "app" / "app.py"
            executable.parent.mkdir(parents=True)
            executable.write_text("#!/bin/bash\n", encoding="utf-8")
            app_file.parent.mkdir(parents=True)
            app_file.write_text("# app\n", encoding="utf-8")

            command, is_handoff = product_master_launch_command(
                app_file, "/runtime/bin/python", platform="darwin"
            )

            self.assertEqual(containing_app_bundle(app_file), bundle)
            self.assertEqual(dedicated_product_master_bundle(app_file), executable.parents[2])
            self.assertFalse(is_handoff)
            self.assertEqual(
                command,
                ["/runtime/bin/python", str(app_file.resolve()), "--product-master"],
            )

    def test_non_macos_source_launches_the_editor_directly(self):
        app_file = Path("/workspace/app.py")
        command, is_handoff = product_master_launch_command(
            app_file, "/runtime/bin/python", platform="linux"
        )
        self.assertFalse(is_handoff)
        self.assertEqual(command, ["/runtime/bin/python", str(app_file.resolve()), "--product-master"])

    def test_macos_without_a_bundle_has_a_safe_direct_fallback(self):
        app_file = Path("/workspace/not-bundled-app.py")
        command, is_handoff = product_master_launch_command(
            app_file, "/runtime/bin/python", platform="darwin"
        )
        self.assertFalse(is_handoff)
        self.assertEqual(command, ["/runtime/bin/python", str(app_file.resolve()), "--product-master"])


if __name__ == "__main__":
    unittest.main()
