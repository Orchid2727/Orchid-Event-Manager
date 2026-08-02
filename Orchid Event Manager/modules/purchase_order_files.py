"""Identify final purchase-order PDFs without hiding valid event worksheets."""

from __future__ import annotations

from pathlib import Path


_DECORATION_AND_SUPPORTING_REPORT_TERMS = (
    "screen print", "screen_print", "screen-print", "embroidery job",
    "embroidery_job", "decoration report", "decoration_report",
    "outsourced job", "outsourced_job", "outsourced decoration", "outsourced_decoration",
    "in-house", "in_house", "receiving and decoration", "receiving_and_decoration",
    "internal_purchase_order", "employee totals", "employee_totals",
    "event summary", "event_summary",
)


def is_final_purchase_order_pdf(path: str | Path) -> bool:
    """Return whether *path* is an orderable PDF in a protected release.

    Uniform Sizing Events intentionally put several boot vendors in one
    ``Ship to Orchid Purchase Order``. It is still the event's final
    purchasing worksheet and
    must remain available on the Purchase Orders page and in Finder.
    """
    name = Path(path).name.casefold()
    if not name.endswith(".pdf") or "purchase_order" not in name:
        return False
    if (
        "ship_to_orchid" in name or "ship to orchid" in name
        or "non-included" in name or "non_included" in name or "non included" in name
    ):
        return True
    return not any(term in name for term in _DECORATION_AND_SUPPORTING_REPORT_TERMS)
