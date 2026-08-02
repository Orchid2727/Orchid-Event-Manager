"""Tru-Spec color-specific purchasing numbers for vendor purchase orders.

Orchid uses a shopper-friendly style number for several Tru-Spec pant families,
but Tru-Spec assigns a different number for each garment color.  The catalog
below is the safe default for the confirmed families.  A separately saved
``Vendor Product #`` remains authoritative, so an intentional Orchid override
is never overwritten.
"""
from __future__ import annotations

import re


def _clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _key(value: object) -> str:
    text = _clean(value).casefold().replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


TRUSPEC_COLOR_PRODUCT_NUMBERS = {
    "24/7 original tactical pant": {
        "khaki": "1060", "navy": "1061", "black": "1062", "coyote": "1063",
        "od green": "1064", "olive drab": "1064", "brown": "1065",
        "charcoal": "1079", "grey": "1089", "gray": "1089",
    },
    "pro flex pant": {
        "khaki": "1482", "black": "1483", "le green": "1484", "navy": "1485",
        "coyote": "1486", "le green black": "1487",
    },
    "agility pant": {
        "khaki": "1524", "navy": "1525", "black": "1526", "ranger green": "1527",
        "flat dark earth": "1528", "fde": "1528",
    },
    "pro vector pant": {
        "khaki": "1555", "coyote": "1556", "le green": "1557", "black": "1560",
        "navy": "1566",
    },
    "24/7 original tactical pant 2.0": {
        "black": "2220", "khaki": "2222", "ranger green": "2224", "fde": "2225",
        "flat dark earth": "2225", "charcoal": "2226", "buckskin": "2227",
        "lapd blue": "2230", "le green": "2231", "light grey": "2232",
        "light gray": "2232", "navy": "2233", "od green": "2234",
        "olive drab": "2234",
    },
}

KNOWN_TRUSPEC_PRODUCT_NUMBERS = {
    number
    for color_map in TRUSPEC_COLOR_PRODUCT_NUMBERS.values()
    for number in color_map.values()
}


def _family(product_name: object) -> str:
    text = _key(product_name)
    # Check 2.0 first because it includes the original-pant wording.
    if "24 7 original tactical pant 2 0" in text:
        return "24/7 original tactical pant 2.0"
    if "pro flex pant" in text:
        return "pro flex pant"
    if "agility pant" in text:
        return "agility pant"
    if "pro vector pant" in text:
        return "pro vector pant"
    if "24 7 original tactical pant" in text or "original tactical pant rip stop" in text:
        return "24/7 original tactical pant"
    return ""


def default_vendor_product_number(
    vendor: object,
    product_name: object,
    garment_color: object,
    existing_number: object = "",
) -> str:
    """Return the confirmed Tru-Spec purchasing number for one garment row.

    ``existing_number`` is a deliberate Vendor Product # override—not Orchid's
    shared Shopify style.  Known but mismatched Tru-Spec values are replaced so
    a stale color row (for example, Pro Vector Coyote stored as 1555) cannot
    leak into a new purchase order.
    """
    existing = _clean(existing_number)
    if _key(vendor) != "tru spec":
        return existing
    family = _family(product_name)
    if not family:
        return existing
    expected = TRUSPEC_COLOR_PRODUCT_NUMBERS[family].get(_key(garment_color), "")
    if existing and (not expected or existing not in KNOWN_TRUSPEC_PRODUCT_NUMBERS or existing == expected):
        return existing
    return expected
