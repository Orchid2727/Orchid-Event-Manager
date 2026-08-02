"""Regression checks for the one-page outsourced-report cover overview."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

import pandas as pd
from pypdf import PdfReader

from modules.final_po_generator import _build_outsourced_decoration_report


class OutsourcedCoverLayoutTests(unittest.TestCase):
    def test_both_report_departments_keep_a_reference_artwork_control(self):
        """Artwork can be managed on either tab before color rows exist."""
        app_source = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
        self.assertIn("def _outsourced_cover_artwork_reference_spec", app_source)
        self.assertIn('"key": f"{selected}::artwork-reference"', app_source)
        self.assertIn("display_specs = [reference_spec, *specs]", app_source)
        self.assertIn('if department else ["embroidery", "screen-printing"]', app_source)
        self.assertIn('text="Replace" if exists else "Upload"', app_source)
        self.assertIn('text="Open"', app_source)
        self.assertIn('text="Remove"', app_source)

    def test_multicolor_count_headers_stay_on_one_line(self):
        """Both count headings remain readable while the cover stays one page."""
        colors = ["Black", "Navy", "Silver", "White"]
        artwork_path = Path(__file__).resolve().parents[1] / "assets" / "orchid_logo.png"
        detail = pd.DataFrame([
            {
                "Style Number": f"CTA-{index}",
                "Product Name": "Performance Cap",
                "Garment Color": "Black",
                "Size": "One Size",
                "Quantity": index + 1,
                "Operational Decoration Type": "Embroidery",
                "Operational Decoration Location": "Left Chest" if index < 3 else "Left Sleeve",
                "Operational Decoration Color": color,
            }
            for index, color in enumerate(colors)
        ])

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "outsourced_embroidery.pdf"
            _build_outsourced_decoration_report(
                detail,
                output,
                "Utilities Department",
                "Entire Order Outsourced",
                outsourced_job_name="Utilities Department",
                cover_artwork=[
                    {
                        "label": f"{color} Thread",
                        "path": artwork_path,
                        "decoration_type": "embroidery",
                    }
                    for color in colors
                ],
                cover_notes="Verify artwork and thread colors before production.",
                department_kind="embroidery",
            )
            reader = PdfReader(str(output))
            cover_text = reader.pages[0].extract_text()

        # A normal four-color job still has one cover page followed by a
        # production page.  The no-split header returns as the complete word.
        self.assertEqual(len(reader.pages), 2)
        # Two headers plus the location and grand-total labels include the
        # complete word.  A split header would instead extract as "PIEC\\nES".
        self.assertEqual(cover_text.count("PIECES"), 4)
        self.assertNotIn("PIEC\nES", cover_text)
        self.assertIn("TOTAL EMBROIDERED PIECES", cover_text)
        self.assertIn("WHITE THREAD", cover_text)
        self.assertIn("OUTSOURCED TO", cover_text)
        self.assertIn("Stitch N Print", cover_text)

    def test_cover_prints_the_department_specific_outsourced_destination(self):
        """A saved non-default destination is prominent on the matching cover."""
        detail = pd.DataFrame([
            {
                "Style Number": "PC54",
                "Product Name": "Essential Tee",
                "Garment Color": "Navy",
                "Size": "Large",
                "Quantity": 18,
                "Operational Decoration Type": "Screen Printing",
                "Operational Decoration Location": "Full Front",
                "Operational Decoration Color": "White",
            },
        ])

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "outsourced_screen_printing.pdf"
            _build_outsourced_decoration_report(
                detail,
                output,
                "Utilities Department",
                "Entire Order Outsourced",
                outsourced_job_name="Utilities Department",
                outsourced_to="Precision Screen Print",
                department_kind="screen-printing",
            )
            reader = PdfReader(str(output))
            cover_text = reader.pages[0].extract_text()

        self.assertIn("OUTSOURCED TO", cover_text)
        self.assertIn("Precision Screen Print", cover_text)
        self.assertNotIn("Stitch N Print", cover_text)


if __name__ == "__main__":
    unittest.main()
