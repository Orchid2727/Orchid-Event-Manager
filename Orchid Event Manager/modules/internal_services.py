from __future__ import annotations

import re
from typing import Any

SEW_ON_PATCH_LABEL = "Sew On Patch (In-House)"
HEMMING_ALTERATION_LABEL = "Hemming / Alteration (In-House)"

IN_HOUSE_DECORATION_TYPES = {
    SEW_ON_PATCH_LABEL.casefold(),
    HEMMING_ALTERATION_LABEL.casefold(),
}

# Orchid product numbers reserved for service/charge lines rather than garments.
# These rows remain available for order totals but must never enter purchasing,
# Purchase Review, vendor POs, or outsourced decoration reports.
INTERNAL_SERVICE_STYLE_CODES = {"750"}

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

# Decoration-only Shopify products and option values. These are accounting or
# production charges, not garments to purchase. Patterns are intentionally
# specific so a normal garment description that merely mentions a logo is not
# excluded.
DECORATION_CHARGE_PATTERNS = (
    r"^(?:one|additional|second|extra|60[-\s]*year)\s+logo(?:\s+(?:charge|fee|service))?$",
    r"^(?:add\s+)?(?:embroidered|embroidery|screen[- ]?print(?:ed|ing)?)\s+logo(?:s)?(?:\s+(?:charge|fee|service))?$",
    r"^(?:one\s+color\s+)?screen[- ]?print(?:ed|ing)?\s+logo(?:s)?(?:\s+(?:charge|fee|service))?$",
    r"^(?:bag|backpack)\s+embroidery(?:\s+(?:charge|fee|service))?$",
    r"^higher\s+stitch\s+count(?:\s+left\s+chest)?\s+logo(?:\s+(?:charge|fee|service))?$",
    r"^(?:larger\s+)?(?:left\s+chest|left\s+sleeve|hat|bag|backpack)\s+logo(?:\s+(?:charge|fee|service))?$",
    r"^(?:add\s+)?(?:embroidery|screen[- ]?print(?:ing)?)(?:\s+(?:charge|fee|service))?$",
    r"^(?:logo|embroidery|screen[- ]?print(?:ing)?)\s+(?:charge|fee|service)$",
    r"^(?:add\s+)?(?:name|title|id|monogram)\s+(?:embroidery|charge|fee|service)$",
)


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def normalize_style_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", clean(value).upper())


def is_internal_service_style(value: Any) -> bool:
    return normalize_style_code(value) in INTERNAL_SERVICE_STYLE_CODES


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


def is_decoration_charge(*values: Any) -> bool:
    """Return True for a decoration-only charge in any product/variant field."""
    for value in values:
        candidate = clean(value).casefold().strip(" -/:,|.")
        if not candidate:
            continue
        if any(re.fullmatch(pattern, candidate, flags=re.IGNORECASE) for pattern in DECORATION_CHARGE_PATTERNS):
            return True
    return False


def decoration_charge_label(*values: Any) -> str:
    """Return the most useful visible label for an excluded decoration charge."""
    for value in values:
        candidate = clean(value)
        if candidate and is_decoration_charge(candidate):
            return candidate
    return clean(values[0]) if values else "Decoration Charge"


def is_in_house_service_product(
    product_name: Any = "",
    original_line: Any = "",
    decoration_type: Any = "",
    style_number: Any = "",
    garment_color: Any = "",
) -> bool:
    """True only for a service-only line, not a normal garment with in-house work.

    A normal hat or garment can use an in-house decoration type and still needs to
    be ordered from its purchase vendor. A line whose own product description is a
    patch/hemming service should be excluded from vendor purchase orders entirely.
    """
    if is_internal_service_style(style_number):
        return True
    if is_decoration_charge(product_name, original_line, garment_color):
        return True
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
