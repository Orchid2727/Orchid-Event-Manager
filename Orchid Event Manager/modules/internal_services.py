from __future__ import annotations

import re
from typing import Any

SEW_ON_PATCH_LABEL = "Sew On Patch (In-House)"
HEMMING_ALTERATION_LABEL = "Hemming / Alteration (In-House)"

IN_HOUSE_DECORATION_TYPES = {
    SEW_ON_PATCH_LABEL.casefold(),
    HEMMING_ALTERATION_LABEL.casefold(),
}

PATCH_SERVICE_PATTERNS = (
    r"\bsew[- ]?on\s+patch(?:es)?\b",
    r"\bpatch(?:es)?\s+on\b",
    r"\bflag\s+patch(?:es)?\b",
    r"\bpatch\s+(?:application|service|fee)\b",
)
HEMMING_SERVICE_PATTERNS = (
    r"\bhemm?ing\b",
    r"\bhem\s+(?:service|fee|pants?|jeans?|trousers?|shorts?)\b",
    r"\balteration(?:s)?\b",
    r"\binseam\s+(?:alteration|adjustment|service)\b",
)


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def is_in_house_decoration(value: Any) -> bool:
    return clean(value).casefold() in IN_HOUSE_DECORATION_TYPES


def is_sew_on_patch(value: Any) -> bool:
    return clean(value).casefold() == SEW_ON_PATCH_LABEL.casefold()


def is_hemming_alteration(value: Any) -> bool:
    return clean(value).casefold() == HEMMING_ALTERATION_LABEL.casefold()


def infer_in_house_decoration(product_name: Any = "", original_line: Any = "") -> str:
    text = " ".join(filter(None, [clean(product_name), clean(original_line)])).casefold()
    if any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in PATCH_SERVICE_PATTERNS):
        return SEW_ON_PATCH_LABEL
    if any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in HEMMING_SERVICE_PATTERNS):
        return HEMMING_ALTERATION_LABEL
    return ""


def is_in_house_service_product(
    product_name: Any = "",
    original_line: Any = "",
    decoration_type: Any = "",
) -> bool:
    """True only for a service-only line, not a normal garment with in-house work.

    A normal hat or garment can use an in-house decoration type and still needs to
    be ordered from its purchase vendor. A line whose own product description is a
    patch/hemming service should be excluded from vendor purchase orders entirely.
    """
    inferred = infer_in_house_decoration(product_name, original_line)
    if inferred:
        return True
    # Require descriptive service wording before treating a line as service-only.
    # The decoration type by itself is not enough because it may belong to a real
    # garment that still must be purchased.
    return False


def vendor_po_decoration_type(decoration_type: Any) -> str:
    """In-house decoration on a real garment is purchased as a blank garment."""
    if is_in_house_decoration(decoration_type):
        from modules.blank_garment_rules import BLANK_DECORATION_LABEL
        return BLANK_DECORATION_LABEL
    return clean(decoration_type)
