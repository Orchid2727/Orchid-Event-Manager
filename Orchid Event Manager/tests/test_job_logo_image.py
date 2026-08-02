"""Regression checks for clean Current Event and report artwork previews."""

from __future__ import annotations

import unittest

from PIL import Image as PILImage, ImageDraw

from modules.job_logo_image import prepare_artwork_for_preview


class JobLogoImageTests(unittest.TestCase):
    def test_colored_outer_canvas_becomes_transparent_and_is_trimmed(self):
        # Simulates a logo pasted on the gray-blue rectangular canvas shown on
        # the Current Event page. The central purple mark must remain intact.
        source = PILImage.new("RGBA", (160, 100), (146, 160, 178, 255))
        draw = ImageDraw.Draw(source)
        draw.ellipse((58, 24, 104, 76), fill=(89, 24, 201, 255))

        prepared, changed = prepare_artwork_for_preview(source)

        self.assertTrue(changed)
        self.assertEqual(prepared.mode, "RGBA")
        self.assertLess(prepared.width, source.width)
        self.assertLess(prepared.height, source.height)
        self.assertEqual(prepared.getpixel((0, 0))[3], 0)
        self.assertEqual(prepared.getpixel((prepared.width // 2, prepared.height // 2))[:3], (89, 24, 201))


if __name__ == "__main__":
    unittest.main()
