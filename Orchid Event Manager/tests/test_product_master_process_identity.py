from __future__ import annotations

import os
import unittest
from unittest.mock import patch

from modules.process_identity import product_master_process_matches


class ProductMasterProcessIdentityTests(unittest.TestCase):
    def test_accepts_the_dedicated_product_master_process(self):
        with patch("modules.process_identity.process_is_running", return_value=True):
            self.assertTrue(
                product_master_process_matches(
                    4123,
                    lambda _pid: "/runtime/bin/python app.py --product-master",
                )
            )

    def test_accepts_the_macos_bundle_editor_process(self):
        with patch("modules.process_identity.process_is_running", return_value=True):
            self.assertTrue(
                product_master_process_matches(
                    4123,
                    lambda _pid: "/Applications/Orchid Purchase Manager.app/Contents/MacOS/OrchidPurchaseManager --product-master",
                )
            )

    def test_rejects_a_reused_pid_for_another_process(self):
        with patch("modules.process_identity.process_is_running", return_value=True):
            self.assertFalse(
                product_master_process_matches(
                    4123,
                    lambda _pid: "/Applications/Other App.app/Contents/MacOS/Other App",
                )
            )

    def test_rejects_a_dead_pid_without_reading_a_command(self):
        called = False

        def reader(_pid):
            nonlocal called
            called = True
            return "--product-master"

        with patch("modules.process_identity.process_is_running", return_value=False):
            self.assertFalse(product_master_process_matches(4123, reader))
        self.assertFalse(called)

    def test_current_process_is_detected_as_running(self):
        with patch("modules.process_identity.os.kill") as kill:
            from modules.process_identity import process_is_running

            self.assertTrue(process_is_running(os.getpid()))
        kill.assert_called_once_with(os.getpid(), 0)


if __name__ == "__main__":
    unittest.main()
