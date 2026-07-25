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


def _has(text: str, pattern: str) -> bool:
    return bool(re.search(pattern, text, flags=re.IGNORECASE))


def infer_category(product_name: Any, style_number: Any = "") -> str:
    text = f"{clean(product_name)} {clean(style_number)}".casefold()
    lower_body_text = re.sub(r"\bshort[- ]sleeve\b", "sleeve", text)

    if _has(text, r"\b(fr|fr[- ]?rated|flame[- ]?resistant|fire[- ]?resistant|flame[- ]?retardant|fire[- ]?retardant)\b"):
        return "Flame Resistant"
    if is_blank_garment_product(product_name, style_number=style_number):
        return "Pants / Jeans / Shorts"
    if _has(lower_body_text, r"\b(jean|jeans|pant|pants|trouser|trousers|shorts|slacks?|cargo pant|bib overall)\b"):
        return "Pants / Jeans / Shorts"
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
            if not clean(result.at[index, column]):
                result.at[index, column] = value
        requires_decoration = normalize_bool(result.at[index, "Requires Decoration"], True)
        decoration_type = clean(result.at[index, "Decoration Type"] if "Decoration Type" in result.columns else "")
        if not requires_decoration and not decoration_type and "Decoration Type" in result.columns:
            result.at[index, "Decoration Type"] = BLANK_DECORATION_LABEL
        if not requires_decoration and "Decoration Color" in result.columns:
            result.at[index, "Decoration Color"] = ""
        if not requires_decoration and "Decoration Location" in result.columns:
            result.at[index, "Decoration Location"] = ""
        if not requires_decoration and "Decoration Placement Instructions" in result.columns:
            result.at[index, "Decoration Placement Instructions"] = ""
    return result
