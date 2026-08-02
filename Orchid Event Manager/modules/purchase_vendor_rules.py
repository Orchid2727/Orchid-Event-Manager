"""Rules that distinguish Orchid's receiving location from a purchase supplier."""

from __future__ import annotations

import re


# Orchid is the receiving business for an internal/non-included order.  It is
# never the supplier that a garment or boot should be ordered from.
_INTERNAL_VENDOR_NAMES = {
    "orchid",
    "orchid uniforms apparel",
    "orchid uniforms and apparel",
}


def _vendor_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).strip()


def is_internal_purchase_vendor(value: object) -> bool:
    """True when a saved vendor is Orchid's internal receiving name.

    These values used to make a report section look like an Orchid purchase
    order even though the actual supplier had not been chosen.  They must
    remain visible for correction, but can never be emitted as a supplier PO.
    """
    return _vendor_key(value) in _INTERNAL_VENDOR_NAMES


INTERNAL_VENDOR_REVIEW_REASON = (
    "Purchase vendor is Orchid Uniforms & Apparel — select the actual supplier"
)
