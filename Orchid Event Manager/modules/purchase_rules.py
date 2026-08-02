from __future__ import annotations

import re
from typing import Any

import pandas as pd

from modules.blank_garment_rules import (
    BLANK_DECORATION_LABEL,
    is_blank_decoration,
    is_blank_garment_product,
)
from modules.internal_services import (
    is_in_house_decoration, is_in_house_service_product, is_internal_service_style,
)

RULE_COLUMNS = [
    "Product Category",
    "Requires Size",
    "Requires Color",
    "Requires Decoration",
]

CATEGORY_OPTIONS = [
    "Other Apparel",
    "T-Shirts",
    "Polos / Shirts",
    "Outerwear",
    "Pants / Jeans / Shorts",
    "Hats / Headwear",
    "Bags",
    "Boots",
    "Drinkware",
    "Towels / Blankets",
    "Safety Workwear",
    "Flame Resistant",
    "Accessories",
    "Other",
]

YES = "Yes"
NO = "No"

# Orchid-specific styles whose requirements are known and must override stale
# Product Master values. 6572 is a one-size embroidered hat; 1104 is a
# Tru-Spec pant and therefore never requires decoration.
ONE_SIZE_HEADWEAR_STYLES = {"6572"}
KNOWN_PANT_STYLES = {"1104", "CT102804"}
KNOWN_COLOR_OPTIONAL_BLANK_STYLES = {"966"}

# These footbeds do not always include the word "insole" in the Shopify title
# (for example, Rocky Air-Port Footbed).  Style number is therefore the
# permanent source of truth.  Treat them as Boots even when a legacy Product
# Master row says Other or still has old color/decoration requirements.
KNOWN_INSOLE_STYLES = {
    "RKK0317",
    "A1Q82",
    "502440",
    "RKK0490",
}


def clean(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_bool(value: Any, default: bool = True) -> bool:
    text = clean(value).casefold()
    if text in {"yes", "y", "true", "1", "required"}:
        return True
    if text in {"no", "n", "false", "0", "optional", "not required"}:
        return False
    return bool(default)


def bool_text(value: bool) -> str:
    return YES if value else NO


def is_known_insole_style(value: Any) -> bool:
    """Return True for a permanent Orchid insole/footbed style."""
    return re.sub(r"\s+", "", clean(value)).upper() in KNOWN_INSOLE_STYLES


def _has(text: str, pattern: str) -> bool:
    return bool(re.search(pattern, text, flags=re.IGNORECASE))


def infer_category(product_name: Any, style_number: Any = "") -> str:
    text = f"{clean(product_name)} {clean(style_number)}".casefold()
    lower_body_text = re.sub(r"\bshort[- ]sleeve\b", "sleeve", text)

    if is_known_insole_style(style_number):
        return "Boots"
    if _has(text, r"\b(fr|fr[- ]?rated|flame[- ]?resistant|fire[- ]?resistant|flame[- ]?retardant|fire[- ]?retardant)\b"):
        return "Flame Resistant"
    if is_blank_garment_product(product_name, style_number=style_number):
        return "Pants / Jeans / Shorts"
    if _has(lower_body_text, r"\b(jean|jeans|pant|pants|trouser|trousers|shorts|slacks?|cargo pant|bib overall)\b"):
        return "Pants / Jeans / Shorts"
    if _has(text, r"\b(boot|boots|insole|insoles)\b"):
        return "Boots"
    if _has(text, r"\b(hat|cap|beanie|boonie|booney|visor|headwear|straw hat)\b"):
        return "Hats / Headwear"
    if _has(text, r"\b(backpack|duffel|tote|messenger|bag|briefcase|cooler)\b"):
        return "Bags"
    if _has(text, r"\b(tumbler|mug|cup|bottle|drinkware|can cooler|koozie)\b"):
        return "Drinkware"
    if _has(text, r"\b(towel|blanket|throw)\b"):
        return "Towels / Blankets"
    if _has(text, r"\b(hi[- ]?vis|high visibility|safety vest|enhanced visibility|reflective)\b"):
        return "Safety Workwear"
    if _has(text, r"\b(jacket|coat|parka|soft shell|softshell|windbreaker|vest|hoodie|fleece|pullover|quarter zip|1/4 zip|full[- ]?zip)\b"):
        return "Outerwear"
    if _has(text, r"\b(polo|button[- ]?down|woven|work shirt|dress shirt|blouse)\b"):
        return "Polos / Shirts"
    if _has(text, r"\b(t[- ]?shirt|tee|long sleeve tee|short sleeve tee)\b"):
        return "T-Shirts"
    if _has(text, r"\b(apron|glove|scarf|sock|belt|accessory|accessories)\b"):
        return "Accessories"
    if _has(text, r"\b(shirt|coverall|uniform|jersey|sweatshirt|crewneck)\b"):
        return "Other Apparel"
    return "Other"


def category_defaults(category: Any, product_name: Any = "", style_number: Any = "") -> dict[str, bool]:
    if is_internal_service_style(style_number) or is_in_house_service_product(
        product_name, decoration_type="", style_number=style_number
    ):
        return {"requires_size": False, "requires_color": False, "requires_decoration": False}
    category_text = clean(category) or infer_category(product_name, style_number)
    if category_text == "Pants / Jeans / Shorts":
        return {"requires_size": True, "requires_color": True, "requires_decoration": False}
    if category_text == "Flame Resistant":
        return {"requires_size": True, "requires_color": True, "requires_decoration": False}
    if category_text == "Boots":
        # Boots (including insoles) are always purchased by size.  They are
        # not decorated, and a separate purchasing color is never required.
        return {"requires_size": True, "requires_color": False, "requires_decoration": False}
    if category_text in {"Hats / Headwear", "Bags", "Drinkware", "Towels / Blankets", "Accessories"}:
        return {"requires_size": False, "requires_color": True, "requires_decoration": True}
    if category_text == "Other":
        # Unknown products remain conservative until categorized. This prevents
        # an incomplete manual item from being treated as order-ready by mistake.
        return {"requires_size": True, "requires_color": True, "requires_decoration": True}
    return {"requires_size": True, "requires_color": True, "requires_decoration": True}


def row_rules(row: pd.Series | dict[str, Any]) -> dict[str, Any]:
    get = row.get
    product_name = clean(get("Product Name", "") or get("Description", ""))
    style_number = clean(get("Style Number", "") or get("Product #", ""))
    style_key = re.sub(r"\s+", "", style_number).upper()

    # These two long-running exceptions are business facts, not suggestions.
    # They deliberately override stale values saved by older versions.
    if is_known_insole_style(style_key):
        return {
            "Product Category": "Boots",
            "Requires Size": YES,
            "Requires Color": NO,
            "Requires Decoration": NO,
        }
    if style_key in ONE_SIZE_HEADWEAR_STYLES:
        category = "Hats / Headwear"
        return {
            "Product Category": category,
            "Requires Size": NO,
            "Requires Color": YES,
            "Requires Decoration": YES,
        }
    if style_key in KNOWN_PANT_STYLES:
        category = "Pants / Jeans / Shorts"
        return {
            "Product Category": category,
            "Requires Size": YES,
            "Requires Color": YES,
            "Requires Decoration": NO,
        }
    if style_key in KNOWN_COLOR_OPTIONAL_BLANK_STYLES:
        category = "Pants / Jeans / Shorts"
        return {
            "Product Category": category,
            "Requires Size": YES,
            "Requires Color": NO,
            "Requires Decoration": NO,
        }

    category = clean(get("Product Category", "")) or infer_category(product_name, style_number)
    defaults = category_defaults(category, product_name, style_number)

    # This is a permanent Product Master policy, not merely a first-time
    # default.  It corrects older boot records that were saved before Boots was
    # introduced or before its purchasing rule was finalized.
    if category.casefold() == "boots":
        return {
            "Product Category": "Boots",
            "Requires Size": YES,
            "Requires Color": NO,
            "Requires Decoration": NO,
        }

    decoration_type = clean(get("Decoration Type", ""))
    service_only = is_internal_service_style(style_number) or is_in_house_service_product(
        product_name, get("Original Line Item", ""), decoration_type, style_number=style_number
    )
    if service_only:
        return {
            "Product Category": category,
            "Requires Size": NO,
            "Requires Color": NO,
            "Requires Decoration": NO,
        }
    requires_decoration = normalize_bool(get("Requires Decoration", ""), defaults["requires_decoration"])
    if decoration_type:
        requires_decoration = not is_blank_decoration(decoration_type) and not is_in_house_decoration(decoration_type)

    return {
        "Product Category": category,
        "Requires Size": bool_text(normalize_bool(get("Requires Size", ""), defaults["requires_size"])),
        "Requires Color": bool_text(normalize_bool(get("Requires Color", ""), defaults["requires_color"])),
        "Requires Decoration": bool_text(requires_decoration),
    }


def apply_purchase_rule_defaults(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    for column in RULE_COLUMNS:
        if column not in result.columns:
            result[column] = ""
    if result.empty:
        return result

    for index, row in result.iterrows():
        rules = row_rules(row)
        for column, value in rules.items():
            # Boots are a locked purchasing category.  Refresh every saved
            # record so an old Color/Decoration setting cannot create a false
            # Product Master or Purchase Review prompt.
            if rules["Product Category"] == "Boots" or not clean(result.at[index, column]):
                result.at[index, column] = value
        requires_decoration = normalize_bool(result.at[index, "Requires Decoration"], True)
        decoration_type = clean(result.at[index, "Decoration Type"] if "Decoration Type" in result.columns else "")
        if (
            not requires_decoration
            and "Decoration Type" in result.columns
            and (rules["Product Category"] == "Boots" or not decoration_type)
        ):
            result.at[index, "Decoration Type"] = BLANK_DECORATION_LABEL
        if not requires_decoration and "Decoration Color" in result.columns:
            result.at[index, "Decoration Color"] = ""
        if not requires_decoration and "Decoration Location" in result.columns:
            result.at[index, "Decoration Location"] = ""
        if not requires_decoration and "Decoration Placement Instructions" in result.columns:
            result.at[index, "Decoration Placement Instructions"] = ""
    return result
