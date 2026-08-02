from __future__ import annotations

import re
from typing import Any

SEW_ON_PATCH_LABEL = "Sew On Patch (In-House)"
HEMMING_ALTERATION_LABEL = "Hemming / Alteration (In-House)"

IN_HOUSE_DECORATION_TYPES = {
    SEW_ON_PATCH_LABEL.casefold(),
    HEMMING_ALTERATION_LABEL.casefold(),
}

# Orchid product numbers reserved for service/charge lines or permanent in-house
# stock rather than vendor-purchased garments. These rows remain available for
# order totals but must never enter Product Master, Purchase Review, vendor POs,
# or outsourced decoration reports. Keep this exact-code list intentionally
# narrow so a normal item that merely mentions a flag, logo, bag, or embroidery
# is never excluded.
INTERNAL_SERVICE_STYLE_CODES = {
    "750",
    # 750B is the Shopify add-on used to charge for a higher-stitch-count
    # embroidered left-chest logo.  It is a charge only, never a garment to
    # purchase or a Product Master record to set up.
    "750B",
    "BAGDECORATION",
    # EMBARK is the one-logo decoration charge in Orchid's Shopify catalog.
    "EMBARK",
    "EMBARK2LOGOS",
    "FLAG",
    "AMERICANFLAG",
    "PARKSLEAFLOGOONLY",
    # Product 800 is a belt held as Orchid bulk inventory. It is sold through
    # Shopify but never needs a vendor purchase or Product Master setup.
    "800",
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

# Decoration-only Shopify products and option values. These are accounting or
# production charges, not garments to purchase. Patterns are intentionally
# specific so a normal garment description that merely mentions a logo is not
# excluded.
DECORATION_CHARGE_PATTERNS = (
    r"^(?:one|additional|second|extra|60[-\s]*year)\s+logo(?:\s+(?:charge|fee|service))?$",
    r"^(?:add\s+)?(?:embroidered|embroidery|screen[- ]?print(?:ed|ing)?)\s+logo(?:s)?(?:\s+(?:charge|fee|service))?$",
    r"^(?:one\s+color\s+)?screen[- ]?print(?:ed|ing)?\s+logo(?:s)?(?:\s+(?:charge|fee|service))?$",
    # Shopify sometimes exports Orchid's Bag Decoration checkout charge with
    # no SKU.  Match its exact title so it remains a sales-total line but is
    # never mistaken for a bag that must be purchased or decorated again.
    r"^bag\s+decoration(?:\s+(?:charge|fee|service))?$",
    r"^(?:bag|backpack)\s+embroidery(?:\s+(?:charge|fee|service))?$",
    r"^higher\s+stitch\s+count(?:\s+embroidered)?(?:\s+left\s+chest)?\s+logo(?:\s+(?:charge|fee|service))?$",
    r"^(?:larger\s+)?(?:left\s+chest|left\s+sleeve|hat|bag|backpack)\s+logo(?:\s+(?:charge|fee|service))?$",
    r"^(?:add\s+)?(?:embroidery|screen[- ]?print(?:ing)?)(?:\s+(?:charge|fee|service))?$",
    r"^(?:logo|embroidery|screen[- ]?print(?:ing)?)\s+(?:charge|fee|service)$",
    r"^(?:add\s+)?(?:name|title|id|monogram)\s+(?:embroidery|charge|fee|service)$",
    # Exact Orchid checkout products. Do not broaden these to generic words:
    # a real garment can legitimately contain the words logo, flag, or Embark.
    r"^embark$",
    r"^embark\s+2\s+logo(?:s)?(?:\s+embroidery)?$",
    r"^american\s+flag(?:\s+(?:logo|embroidery))?$",
    # Parks Leaf Logo is an Orchid checkout charge for decoration only. It has
    # appeared in exports without its service SKU, so recognize the exact
    # product title as well as the PARKSLEAFLOGOONLY code above.
    r"^parks\s+leaf\s+logo(?:\s+only)?(?:\s+(?:charge|fee|service))?$",
)

# A Shopify option occasionally reached Product Master as if it were a product
# (for example, "Medium Black"). It has no product identifier or purchasing
# meaning. This intentionally handles only a bare size-plus-color label, not a
# normal description that happens to mention either word.
_PHANTOM_SIZE_WORDS = {"xxs", "xs", "s", "small", "m", "medium", "l", "large", "xl", "xxl", "2xl", "3xl", "4xl", "5xl", "6xl"}
_PHANTOM_COLOR_WORDS = {"black", "white", "navy", "blue", "royal", "red", "maroon", "green", "olive", "khaki", "gray", "grey", "silver", "brown", "tan", "coyote", "orange", "yellow", "purple", "pink", "charcoal", "gold"}

# "Final Sale Boot" is Orchid's intentionally generic Shopify checkout item
# for a clearance boot already in stock.  It records the $50 sale, but there is
# no product to buy, decorate, or set up in Product Master.  Keep this match
# deliberately exact so a real boot that happens to be on final sale still
# follows the normal Boots purchasing workflow.
FINAL_SALE_ACCOUNTING_PRODUCT_PATTERNS = (
    # Shopify/Report Toaster may append the selected shoe size to the product
    # title (for example, ``Final Sale Boot - 12``).
    r"^final\s+sale\s+boots?(?:\s*[-–—:]\s*.+)?$",
)


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def normalize_style_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", clean(value).upper())


def is_internal_service_style(value: Any) -> bool:
    return normalize_style_code(value) in INTERNAL_SERVICE_STYLE_CODES


def is_phantom_product_label(value: Any) -> bool:
    """Return True only for a bare Shopify size/color value posing as a product."""
    words = re.findall(r"[a-z0-9]+", clean(value).casefold())
    word_pair = (
        len(words) == 2
        and any(word in _PHANTOM_SIZE_WORDS for word in words)
        and any(word in _PHANTOM_COLOR_WORDS for word in words)
    )
    if word_pair:
        return True
    # Product-number normalization removes spaces, so protect the same value
    # when it reaches the candidate synchronizer as MEDIUMBLACK.
    compact = "".join(words)
    return any(
        compact in {f"{size}{color}", f"{color}{size}"}
        for size in _PHANTOM_SIZE_WORDS
        for color in _PHANTOM_COLOR_WORDS
    )


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


def is_final_sale_accounting_product(*values: Any) -> bool:
    """Return True for Orchid's exact no-purchase clearance-boot line."""
    for value in values:
        candidate = clean(value).casefold().strip(" -/:,|.")
        if candidate and any(
            re.fullmatch(pattern, candidate, flags=re.IGNORECASE)
            for pattern in FINAL_SALE_ACCOUNTING_PRODUCT_PATTERNS
        ):
            return True
    return False


def decoration_charge_label(*values: Any) -> str:
    """Return the most useful visible label for an excluded accounting line."""
    for value in values:
        candidate = clean(value)
        if candidate and (
            is_decoration_charge(candidate)
            or is_final_sale_accounting_product(candidate)
        ):
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
    if is_phantom_product_label(product_name) or is_phantom_product_label(original_line):
        return True
    if is_final_sale_accounting_product(product_name, original_line, garment_color):
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
