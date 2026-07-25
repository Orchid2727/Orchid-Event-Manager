"""Shared Page X of Y footer for Orchid's generated PDF reports."""
from __future__ import annotations

from reportlab.lib import colors
from reportlab.pdfgen import canvas


PAGE_NUMBER_COLOR = colors.HexColor("#6F667A")


class PageNumberCanvas(canvas.Canvas):
    """Render a centered total-page footer after ReportLab knows every page."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: list[dict] = []

    def showPage(self):  # noqa: N802 - ReportLab's canvas API uses this name.
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._saved_page_states)
        for page_state in self._saved_page_states:
            self.__dict__.update(page_state)
            self._draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_page_number(self, page_count: int) -> None:
        page_width, _page_height = self._pagesize
        self.saveState()
        self.setFont("Helvetica", 7.5)
        self.setFillColor(PAGE_NUMBER_COLOR)
        self.drawCentredString(
            page_width / 2,
            14,
            f"Page {self._pageNumber} of {page_count}",
        )
        self.restoreState()
